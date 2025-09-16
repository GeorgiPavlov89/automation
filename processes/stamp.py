from __future__ import annotations
import os, sys, argparse, io
from pathlib import Path
from datetime import datetime

import fitz
from PIL import Image, ImageDraw, ImageFont

DEFAULT_ANCHOR_REL = (0.97, 0.02)      
DEFAULT_BORDER_MM  = 0.3                 
DEFAULT_PADDING_MM = 0.8                 
DEFAULT_FONT_SIZE  = 10                 
DEFAULT_AS_IMAGE   = True                
DEFAULT_FILL_BG    = True                
DEFAULT_STROKE_A   = 1.0
DEFAULT_FILL_A     = 1.0
DEFAULT_REL_FALLBACK = (0.76, 0.06, 0.97, 0.16)  # ако auto-fit е изключен

# ---------------- Desktop detection ----------------
def _desktop_from_onedrive() -> Path | None:
    od = os.environ.get("OneDrive")
    p = Path(od) / "Desktop" if od else None
    return p if p and p.exists() else None

def _desktop_from_registry() -> Path | None:
    try:
        import winreg
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
        ) as key:
            val, _ = winreg.QueryValueEx(key, "Desktop")
        p = Path(os.path.expandvars(val))
        return p if p.exists() else None
    except Exception:
        return None

def get_desktop_dir() -> Path:
    return _desktop_from_onedrive() or _desktop_from_registry() or (
        Path(os.environ.get("USERPROFILE", str(Path.home()))) / "Desktop"
    )

# ---------------- geometry helpers ----------------
PT_PER_MM = 72 / 25.4
def mm(x: float) -> float: return x * PT_PER_MM

def inset(rect: fitz.Rect, pad: float) -> fitz.Rect:
    return fitz.Rect(rect.x0 + pad, rect.y0 + pad, rect.x1 - pad, rect.y1 - pad)

# ---------------- TXT ----------------
def read_stamp_txt(root: Path) -> tuple[str|None, str|None]:
    """Чете NAME и REG_NO от <root>/stamp.txt (формат: NAME=..., REG_NO=...)."""
    txt = root / "stamp.txt"
    name = reg_no = None
    if txt.exists():
        try:
            with open(txt, encoding="utf-8") as f:
                for line in f:
                    if "=" in line:
                        k, v = line.split("=", 1)
                        k = k.strip().upper(); v = v.strip()
                        if k == "NAME":   name = v
                        if k == "REG_NO": reg_no = v
        except Exception as e:
            print(f"Error reading stamp.txt: {e}")
    return name, reg_no

# ---------------- Excel fallback (CASE) ----------------
def read_first_case_no() -> str | None:
    """Взема първото дело от Reports_Order.* (ако успее)."""
    try:
        from automation.io.excel_reader import read_cases
        rows = read_cases(None)
        return rows[0].get("case_no") if rows else None
    except Exception:
        return None

# ---------------- font helpers ----------------
def _choose_font_path(font_path: Path | None) -> Path:
    candidates = []
    if font_path: candidates.append(Path(font_path))
    candidates += [
        Path(r"C:\Windows\Fonts\consola.ttf"),
        Path(r"C:\Windows\Fonts\arial.ttf"),
        Path(r"C:\Windows\Fonts\segoeui.ttf"),
        Path(r"C:\Windows\Fonts\times.ttf"),
    ]
    for p in candidates:
        if p and p.exists(): return p
    raise FileNotFoundError("Не намерих подходящ TTF в системата.")

# ---------------- text → PNG with auto-sizing ----------------
def measure_and_render_text_png(
    text: str,
    font_path: Path | None,
    font_size_pt: float,
    pad_px: int = 8,
    align_right: bool = True,
):
    """
    Връща (png_bytes, px_w, px_h, scale), като картинката е плътна около текста.
    """
    SCALE = 6.0  # ~432 DPI
    # първо измерване с голямо „платно“
    tmp_w, tmp_h = 2000, 1000
    img = Image.new("RGBA", (tmp_w, tmp_h), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    fpath = _choose_font_path(font_path)
    fnt = ImageFont.truetype(str(fpath), int(font_size_pt * SCALE))

    lines = text.split("\n") or [""]
    # височина на ред и обща височина
    line_h = max(draw.textbbox((0,0), "Ag", font=fnt)[3], 1)
    total_h = int(len(lines) * line_h * 1.10)  # малко по-компактно

    # максимална ширина
    max_w = 0
    for ln in lines:
        w = draw.textbbox((0,0), ln, font=fnt)[2]
        if w > max_w: max_w = w

    W = max_w + 2 * pad_px
    H = total_h + 2 * pad_px

    # второ рендериране върху точен размер
    img2 = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    draw2 = ImageDraw.Draw(img2)

    y = pad_px
    for ln in lines:
        tw = draw2.textbbox((0,0), ln, font=fnt)[2]
        x = W - pad_px - tw if align_right else pad_px
        draw2.text((x, y), ln, font=fnt, fill=(0,0,0,255))
        y += int(line_h * 1.10)

    buf = io.BytesIO()
    img2.save(buf, format="PNG")
    return buf.getvalue(), W, H, SCALE

# ---------------- stamping core ----------------
def stamp_one(pdf_in: Path, out: Path, *,
              # позициониране
              anchor_rel: tuple[float,float] | None,
              rel_fallback: tuple[float,float,float,float] | None,
              margin_mm: float, width_mm: float | None, height_mm: float | None,
              page_index: int,
              # данни
              name: str | None, reg_no: str | None,
              doc_no: str | None, in_date: str | None, case_no: str | None,
              # шрифт
              font_file: Path | None, font_size: float,
              # визия
              as_image: bool,
              border_mm: float, padding_mm: float,
              fill_white: bool, stroke_alpha: float, fill_alpha: float,
              debug_frame: bool=False):

    doc = fitz.open(pdf_in)
    page = doc[page_index]
    Wp, Hp = page.rect.width, page.rect.height

    # Текст (редът е като "желаната" визия: първо ЧСИ, после Вх. документ, дата, дело)
    lines = []
    if name and reg_no: lines.append(f"ЧСИ {name} № {reg_no}")
    if doc_no:          lines.append(f"Вх. документ №  {doc_no}")
    if in_date:         lines.append(f"Входиран: {in_date}")
    if case_no:         lines.append(f"Изп. дело:  {case_no}")
    text = "\n".join(lines) if lines else ""

    # 1) Авто-пасваща кутия: рендерираме PNG и вземаме точните пиксели
    if as_image:
        png, px_w, px_h, SCALE = measure_and_render_text_png(
            text=text,
            font_path=font_file,
            font_size_pt=font_size,
            pad_px=8,
            align_right=True
        )
        # конверсия px → pt
        pt_w = px_w / SCALE
        pt_h = px_h / SCALE
    else:
        # ако е векторен текст, ползваме fallback правоъгълник
        pt_w = mm(width_mm) if width_mm else mm(55.0)
        pt_h = mm(height_mm) if height_mm else pt_w * 0.6
        png = None

    # 2) Позициониране (anchor горе-вдясно)
    if anchor_rel:
        right = Wp * anchor_rel[0]
        top   = Hp * anchor_rel[1]
        rect  = fitz.Rect(right - pt_w, top, right, top + pt_h)
    elif rel_fallback:
        x0 = Wp * rel_fallback[0]; y0 = Hp * rel_fallback[1]
        x1 = Wp * rel_fallback[2]; y1 = Hp * rel_fallback[3]
        rect = fitz.Rect(x0, y0, x1, y1)
    else:
        # последен резерв – горе-вдясно с марджин
        m = mm(margin_mm)
        rect = fitz.Rect(Wp - m - pt_w, m, Wp - m, m + pt_h)

    # 3) Рамка + фон (векторно)
    shape = page.new_shape()
    shape.draw_rect(rect)
    fill = (1,1,1) if fill_white else None
    shape.finish(
        width = max(0.2, mm(border_mm)),
        color = (0,0,0),
        fill  = fill,
        stroke_opacity = max(0.0, min(1.0, stroke_alpha)),
        fill_opacity   = max(0.0, min(1.0, fill_alpha)),
    )
    shape.commit()

    if debug_frame:
        page.draw_rect(rect, color=(1,0,0), width=0.5)

    # 4) Текст вътре
    inner = inset(rect, mm(padding_mm))
    if as_image:
        page.insert_image(inner, stream=png, keep_proportion=False)
    else:
        # векторен текст с подравняване вдясно
        kwargs = {
            "fontsize": font_size,
            "align": fitz.TEXT_ALIGN_RIGHT,
            "color": (0,0,0),
        }
        fpath = font_file if font_file else _choose_font_path(None)
        if fpath and Path(fpath).exists():
            kwargs["fontfile"] = str(fpath)
        page.insert_textbox(inner, text, **kwargs)

    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    doc.close()

# ---------------- main ----------------
def main():
    desktop = get_desktop_dir()
    root = desktop / "Робот-Дела"
    bnb_dir = root / "BNB"
    out_dir = bnb_dir / "Stamped"

    ap = argparse.ArgumentParser(
        description="BNB печат – авто-пасваща рамка горе вдясно.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    # Вход/изход
    ap.add_argument("--in",  dest="in_path",  default=str(bnb_dir), help="Вход: PDF или папка с PDF-и")
    ap.add_argument("--out", dest="out_dir",  default=str(out_dir), help="Изходна папка")
    ap.add_argument("--page", type=int, default=0)

    # Данни
    ap.add_argument("--name", default=None)
    ap.add_argument("--reg",  default=None)
    ap.add_argument("--doc",  default=None)
    ap.add_argument("--date", default=None)
    ap.add_argument("--case", default=None)

    # Шрифт/визия (имат дефолти — не са нужни флагове)
    ap.add_argument("--font", default=None)
    ap.add_argument("--font-size", type=float, default=DEFAULT_FONT_SIZE)
    ap.add_argument("--as-image", action="store_true", default=DEFAULT_AS_IMAGE)

    # Позициониране/кутия
    ap.add_argument("--anchor-rel", nargs=2, type=float, metavar=("RIGHT","TOP"), default=DEFAULT_ANCHOR_REL)
    ap.add_argument("--rel-fallback", nargs=4, type=float, metavar=("X0","Y0","X1","Y1"),
                    default=DEFAULT_REL_FALLBACK)
    ap.add_argument("--margin-mm", type=float, default=5.0)
    ap.add_argument("--w-mm", type=float, default=None)
    ap.add_argument("--h-mm", type=float, default=None)
    ap.add_argument("--border-mm", type=float, default=DEFAULT_BORDER_MM)
    ap.add_argument("--padding-mm", type=float, default=DEFAULT_PADDING_MM)
    ap.add_argument("--stroke-alpha", type=float, default=DEFAULT_STROKE_A)
    ap.add_argument("--fill-alpha",   type=float, default=DEFAULT_FILL_A)
    ap.add_argument("--no-fill", action="store_true", default=(not DEFAULT_FILL_BG))
    ap.add_argument("--debug-frame", action="store_true", default=False)

    args = ap.parse_args()

    # Данни от TXT (CLI има приоритет)
    name, reg_no = read_stamp_txt(root)
    doc_str  = args.doc
    case_str = args.case or read_first_case_no()
    date_str = args.date or datetime.now().strftime("%d.%m.%Y")

    in_path = Path(args.in_path)
    out_dir = Path(args.out_dir)

    # входни PDF-и
    if in_path.is_dir():
        pdfs = sorted([p for p in in_path.glob("*.pdf") if p.is_file()])
    elif in_path.is_file() and in_path.suffix.lower() == ".pdf":
        pdfs = [in_path]
    else:
        print(f"❌ Невалиден вход: {in_path}"); sys.exit(2)
    if not pdfs:
        print(f"ℹ Няма PDF-и в {in_path}"); return

    font_file = Path(args.font) if args.font else None

    for p in pdfs:
        out = out_dir / (p.stem + "_stamped.pdf")
        stamp_one(
            pdf_in=p, out=out,
            anchor_rel=tuple(args.anchor_rel) if args.anchor_rel else None,
            rel_fallback=tuple(args.rel_fallback) if args.rel_fallback else None,
            margin_mm=args.margin_mm, width_mm=args.w_mm, height_mm=args.h_mm,
            page_index=args.page,
            name=args.name or name, reg_no=args.reg or reg_no,
            doc_no=doc_str, in_date=date_str, case_no=case_str,
            font_file=font_file, font_size=args.font_size,
            as_image=args.as_image,
            border_mm=args.border_mm, padding_mm=args.padding_mm,
            fill_white=(not args.no_fill),
            stroke_alpha=args.stroke_alpha, fill_alpha=args.fill_alpha,
            debug_frame=args.debug_frame
        )
        print(f"✅ {p.name} → {out.name}")

if __name__ == "__main__":
    main()
