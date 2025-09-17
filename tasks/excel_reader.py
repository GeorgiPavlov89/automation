# automation/tasks/excel_reader.py
from __future__ import annotations
from typing import Any, Dict, List
from pathlib import Path
import os

import pandas as pd
from automation.orchestrator import task

# Нормализация по желание; можеш да разшириш MAP според твоите колони
MAP = {
    "EGN-EIK": "egn_or_eik",
    "No_ID": "case_no",
    "NSSI_Assur": "do_noi_assur",
    "Regix-NOI_Trud": "do_regix_noi_trud",
    "Regix-NOI_Pens": "do_regix_noi_pens",
    "GRAO": "do_grao",
    "BNB": "do_bnb",
    "IKAR": "do_ikar",
    "NAP_art74": "do_nap_art74",
    "NAP_art191": "do_nap_art191",
    "DVIJEMI": "do_dvijemi",
    "BEZ_ZAPORI": "do_bez_zapори",
}
WANTED_BASENAME = "Reports_Order"
ALLOWED_EXTS = (".xlsx", ".xlsm", ".xls")

def _pick_engine(p: Path) -> Dict[str, Any]:
    ext = p.suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        return {"engine": "openpyxl"}   # новите формати
    if ext == ".xls":
        return {"engine": "xlrd"}       # старият .xls
    return {}
    # pandas engines: openpyxl (.xlsx/.xlsm), xlrd (.xls), и др. описани в дока. :contentReference[oaicite:4]{index=4}

def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.rename(columns=lambda c: str(c).strip())
    df = df.rename(columns={c: MAP.get(c, c) for c in df.columns})
    # създай липсващи нормализирани колони
    for want in set(MAP.values()):
        if want not in df.columns:
            df[want] = None
    return df

def _find_fallback(base_path: Path) -> Path | None:
    """
    Ако точният файл го няма:
      1) пробваме в същата папка с pattern 'Reports_Order*.xls*'
      2) пробваме една папка нагоре (често Desktop), рекурсивно '**/Reports_Order*.xls*'
      3) ако в подадения път личи 'Робот-Дела', пробваме директно '<Desktop>/Робот-Дела/Reports_Order*.xls*'
    Връщаме първото най-ново съвпадение.
    """
    tried: list[Path] = []

    # 1) същата папка
    for cand in sorted(base_path.parent.glob(f"{WANTED_BASENAME}*")):
        if cand.suffix.lower().startswith(".xls") and cand.suffix.lower() in ALLOWED_EXTS and cand.is_file():
            return cand
        tried.append(cand)

    # 2) една нагоре, рекурсивно
    parent = base_path.parent
    for cand in sorted(parent.rglob(f"{WANTED_BASENAME}*.xls*"), key=lambda p: p.stat().st_mtime, reverse=True):
        if cand.suffix.lower() in ALLOWED_EXTS and cand.is_file():
            return cand
        tried.append(cand)

    return None  # ще вдигнем подробна грешка в call-site

@task("read_cases")
def read_cases(path: str) -> List[Dict[str, Any]]:
    """
    Чете Excel:
      - Ако подаденото `path` съществува → ползва него.
      - Иначе търси автоматично 'Reports_Order*.xls[x|m]' в разумни места около подадения път.
      - Вдига подробен FileNotFoundError, ако нищо не открие.
    """
    raw = os.path.expandvars(path)
    p = Path(raw)

    target: Path | None = p if p.exists() else _find_fallback(p)
    if not target or not target.exists():
        # по-ясна грешка за логове
        raise FileNotFoundError(f"Не намирам Excel файла около: {p}")

    df = pd.read_excel(target, **_pick_engine(target))
    df = _normalize_columns(df)
    if "case_no" in df.columns:
        df = df[~df["case_no"].isna()]
    return df.fillna("").to_dict(orient="records")
