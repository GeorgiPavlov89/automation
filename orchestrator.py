from __future__ import annotations
import argparse, json, os, pathlib, logging, logging.handlers, sys
from typing import Callable, Dict, Any, List
import yaml
from importlib import import_module

REGISTRY: dict[str, Callable[..., dict]] = {}

def task(name: str):
    def deco(fn: Callable[..., dict]):
        REGISTRY[name] = fn
        return fn
    return deco

def _package_local_dir() -> pathlib.Path:
    base = pathlib.Path(os.environ.get("LOCALAPPDATA", "")) / "Packages"
    for p in base.glob("*MyAutomation*"):
        return p / "LocalCache" / "MyAutomation"
    return pathlib.Path.home() / "AppData" / "Local" / "MyAutomation"

def _make_logger(name="orchestrator", level=logging.INFO, to_console=True) -> logging.Logger:
    """
    Настройва logging така, че винаги да показва на конзолата и да пише във файл,
    дори ако някой вече е конфигурирал logging преди нас.
    """
    # 1) Нулирай предишните настройки и вдигни root-а (Python 3.8+: force=True)
    logging.basicConfig(level=level, force=True)  # принудителна глобална конфигурация
    # refs: logging HOWTO / basicConfig и StreamHandler. 
    # StreamHandler праща изхода към sys.stdout/sys.stderr. (docs) 

    out_dir = _package_local_dir() / "logs"
    out_dir.mkdir(parents=True, exist_ok=True)
    log_path = out_dir / "worker.log"

    log = logging.getLogger(name)
    log.setLevel(level)
    log.propagate = False  # да не дублира към root

    # --- File handler (rotating) ---
    fh = logging.handlers.RotatingFileHandler(
        log_path, maxBytes=1_000_000, backupCount=5, encoding="utf-8"
    )
    fh.setLevel(level)
    fh.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    log.addHandler(fh)

    # --- Console handler (stdout) ---
    if to_console:
        sh = logging.StreamHandler(stream=sys.stdout)
        sh.setLevel(level)
        sh.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))
        log.addHandler(sh)

    # Мини „банер“, за да си сигурен, че логът е активен
    log.info("Logger ready → file=%s, level=%s", log_path, logging.getLevelName(level))
    return log

def _load_yaml(path: pathlib.Path) -> dict:
    with path.open("r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def _resolve_vars(val: Any, vars: dict, ctx: dict) -> Any:
    if isinstance(val, str):
        try:
            out = val.format(**{**vars, **ctx})
        except Exception:
            out = val
        return os.path.expandvars(out)
    if isinstance(val, dict):
        return {k: _resolve_vars(v, vars, ctx) for k, v in val.items()}
    if isinstance(val, (list, tuple)):
        t = type(val)
        return t(_resolve_vars(v, vars, ctx) for v in val)
    return val

def _run_step(step: dict, ctx: dict, log: logging.Logger):
    name = step["task"]
    mode = step.get("mode", "task")
    result_key = step.get("result_key")
    kwargs_raw = step.get("kwargs", {}) or {}

    vars_dict = ctx.get("__vars__", {})
    kwargs = _resolve_vars(kwargs_raw, vars_dict, ctx)

    mod, attr = name.split(":") if ":" in name else (name, None)
    module = import_module(mod)               # официалният API за динамични импорти
    fn = getattr(module, attr) if attr else module

    log.info("START %s %s", name, kwargs if kwargs else "")
    if mode == "raw":
        out = fn(**kwargs)
        if result_key is not None:
            ctx[result_key] = out
            # кратък helpful log
            if isinstance(out, list):
                log.info("→ %s: %d items", result_key, len(out))
            elif isinstance(out, dict):
                if "stamped_count" in out:
                    log.info("→ stamped_count=%s, output_dir=%s",
                             out.get("stamped_count"), out.get("output_dir"))
                else:
                    log.info("→ %s keys: %s", result_key, ", ".join(sorted(out.keys())))
    else:
        out = fn(dict(ctx), **kwargs)
        if isinstance(out, dict):
            ctx.update(out)
    log.info("END   %s", name)
    return ctx

def _summary_line(ctx: dict) -> str:
    parts = []
    if isinstance(ctx.get("credentials"), list):
        parts.append(f"credentials={len(ctx['credentials'])}")
    if isinstance(ctx.get("cases"), list):
        parts.append(f"cases={len(ctx['cases'])}")
    if "stamped_count" in ctx:
        sc = ctx.get("stamped_count")
        od = ctx.get("output_dir")
        parts.append(f"stamped={sc} -> {od}")
    return " | ".join(parts) if parts else "(no outputs captured)"

def main() -> dict:
    ap = argparse.ArgumentParser(description="Simple task orchestrator")
    ap.add_argument("--config", default="pipelines.yml", help="Path to pipelines.yml")
    ap.add_argument("--verbose", action="store_true", help="Verbose console logging (DEBUG)")
    args = ap.parse_args()

    level = logging.DEBUG if args.verbose else logging.INFO
    log = _make_logger(level=level, to_console=True)

    cfg_path = pathlib.Path(args.config)
    if not cfg_path.exists():
        here = pathlib.Path(__file__).parent
        for cand in [here / args.config, here / "pipelines.yml", here.parent / "pipelines.yml"]:
            if cand.exists():
                cfg_path = cand
                break
    log.info("Using config: %s", cfg_path)

    cfg = _load_yaml(cfg_path)
    vars_cfg = cfg.get("vars", {}) or {}
    selected = cfg.get("use")
    pipeline: List[dict] = cfg["pipelines"][selected]

    ctx: Dict[str, Any] = {"__vars__": vars_cfg}

    for step in pipeline:
        cond = step.get("when")
        if cond and "file_exists" in cond:
            p = pathlib.Path(_resolve_vars(cond["file_exists"], vars_cfg, ctx))
            if not p.exists():
                log.info("SKIP %s (missing %s)", step["task"], p)
                continue
        ctx = _run_step(step, ctx, log)

    safe_ctx = {k: ("***" if "pass" in k.lower() else v) for k, v in ctx.items()}
    log.info("PIPELINE OK; %s", _summary_line(ctx))
    log.debug("context=%s", json.dumps(safe_ctx, ensure_ascii=False))
    return ctx

if __name__ == "__main__":
    main()
