# automation/io/excel_reader.py
from __future__ import annotations
r"""
Чете и нормализира Excel с „дела“. Ако не подадеш път, автоматично ще вземе:
  <Desktop>\Робот-Дела\Reports_Order.xlsx   (или .xls)

Покрива:
- OneDrive Desktop (Known Folder Move);
- пренасочен Desktop на друг диск (чете HKCU\...\User Shell Folders\Desktop).

Зависимости:
    pip install pandas openpyxl xlrd
(За .xlsx/.xlsm е нужен openpyxl; за .xls е нужен xlrd.)
"""

import os
import re
import datetime as dt
from pathlib import Path
from typing import List, Dict
import pandas as pd

# ---------- Автоматично откриване на Desktop ----------

def _desktop_from_onedrive() -> Path | None:
    od = os.environ.get("OneDrive")
    if od:
        p = Path(od) / "Desktop"
        if p.exists():
            return p
    return None

def _desktop_from_registry() -> Path | None:
    r"""Чете HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\Desktop."""
    try:
        import winreg
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
        ) as key:
            value, _typ = winreg.QueryValueEx(key, "Desktop")
            expanded = os.path.expandvars(value)  # %USERPROFILE% / %OneDrive%
            p = Path(expanded)
            if p.exists():
                return p
    except Exception:
        pass
    return None

def _desktop_from_userprofile() -> Path:
    up = os.environ.get("USERPROFILE") or str(Path.home())
    p = Path(up) / "Desktop"
    return p if p.exists() else Path.home() / "Desktop"

def get_desktop_dir() -> Path:
    """Връща реалния Desktop, дори да е пренасочен (OneDrive/Registry)."""
    return (
        _desktop_from_onedrive()
        or _desktop_from_registry()
        or _desktop_from_userprofile()
    )

def resolve_cases_path(
    folder_name: str = "Робот-Дела",
    base_name: str = "Reports_Order"
) -> Path:
    """
    Търси във <Desktop>\<folder_name>\:
      - Reports_Order.xlsx, после Reports_Order.xls.
    Ако липсва, връща очакваното място за .xlsx и създава папката.
    """
    desktop = get_desktop_dir()
    print(f"[excel_reader] Открит Desktop: {desktop}")
    root = desktop / folder_name
    root.mkdir(parents=True, exist_ok=True)

    candidates = [
        root / f"{base_name}.xlsx",
        root / f"{base_name}.xls",
    ]
    for c in candidates:
        if c.exists():
            return c
    return candidates[0]  # очаквано място, ако файлът още не съществува

# ---------- Мап към нормализирани ключове ----------

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
    "BEZ_ZAPORI": "do_bez_zapori",
}
FLAGS: List[str] = [v for v in MAP.values() if v.startswith("do_")]
TEXT_TRUE = {"1", "1.0", "да", "д", "yes", "y", "true", "t", "✓", "x"}

# ---------- ЕГН валидатор (информативен) ----------

_W = [2,4,8,5,10,9,7,3,6]

def _egn_date_ok(s: str) -> bool:
    yy, mm, dd = int(s[0:2]), int(s[2:4]), int(s[4:6])
    if 1 <= mm <= 12:
        century, real_m = 1900, mm
    elif 21 <= mm <= 32:  # 1800–1899
        century, real_m = 1800, mm - 20
    elif 41 <= mm <= 52:  # 2000–2099
        century, real_m = 2000, mm - 40
    else:
        return False
    Y = century + yy
    try:
        dt.date(Y, real_m, dd)
        return True
    except ValueError:
        return False

def is_valid_egn(s: str) -> bool:
    s = re.sub(r"\D", "", str(s or ""))
    if not re.fullmatch(r"\d{10}", s): return False
    if not _egn_date_ok(s): return False
    chk = sum(int(s[i]) * _W[i] for i in range(9)) % 11
    if chk == 10: chk = 0
    return chk == int(s[9])

def classify_id(value: str) -> str:
    t = re.sub(r"\D", "", str(value or ""))
    if len(t) == 10:
        return "EGN" if is_valid_egn(t) else "EGN_10_INVALID"
    if len(t) == 9:  return "EIK9"
    if len(t) == 13: return "EIK13"
    return "UNKNOWN"

# ---------- Основно четене ----------

def read_cases(path: str | None = None) -> List[Dict]:
    """
    Ако path е None, използваме <Desktop>\Робот-Дела\Reports_Order.xlsx|xls.
    Връща списък от dict записи с нормализирани ключове и флагове 0/1.
    """
    p = Path(path) if path else resolve_cases_path()
    if not p.exists():
        raise FileNotFoundError(
            f"Не намирам входния файл. Очаквам го тук:\n  {p}\n"
            f"Сложи там Reports_Order.xlsx (или .xls). Папката е създадена."
        )

    # Избор на engine по разширение
    ext = p.suffix.lower()
    engine = "xlrd" if ext == ".xls" else ("openpyxl" if ext in (".xlsx", ".xlsm") else None)

    print(f"[excel_reader] Чета: {p} (engine={engine})")
    df = pd.read_excel(p, engine=engine)
    print(f"[excel_reader] Колони (източник): {list(df.columns)}")

    # Почисти заглавията и мапни към нормализирани имена
    df = df.rename(columns=lambda c: str(c).strip())
    df = df.rename(columns={c: MAP.get(c, c) for c in df.columns})
    print(f"[excel_reader] Колони (норм.): {list(df.columns)}")

    # Увери се, че всички флаг-колони съществуват
    for f in FLAGS:
        if f not in df.columns:
            df[f] = 0

    # Нормализирай флаговете: текстови „истини“ → 1, после към int
    for f in FLAGS:
        df[f] = df[f].apply(lambda v: 1 if str(v).strip().lower() in TEXT_TRUE else v)
    df[FLAGS] = df[FLAGS].apply(pd.to_numeric, errors="coerce").fillna(0).astype(int)

    # Класификация/валидация на идентификатора (само информативно)
    df["id_kind"] = df["egn_or_eik"].map(classify_id)
    bad_rows = df.index[df["id_kind"] == "EGN_10_INVALID"]
    if len(bad_rows):
        rows = [int(i) + 2 for i in bad_rows.to_list()]  # +2 за header + 1-index
        print("⚠ ЕГН проверка (10-цифрени, но с грешка) в редове:", rows)

    return df.fillna("").to_dict(orient="records")

# ---------- CLI ----------

if __name__ == "__main__":
    import argparse
    from pprint import pprint
    import sys, traceback

    ap = argparse.ArgumentParser(
        description="Прочети и нормализирай Excel с дела",
        epilog=r"Ако не подадеш път, ще се ползва <Desktop>\Робот-Дела\Reports_Order.*"
    )
    ap.add_argument("path", nargs="?", help=r"(по желание) път до .xls/.xlsx")
    ap.add_argument("--out", default=None, help="(по желание) запиши нормализираните данни в .xlsx")
    args = ap.parse_args()

    try:
        records = read_cases(args.path)
        print(f"\n✅ Прочетени редове: {len(records)}")
        for i, rec in enumerate(records, 1):
            print(f"\n--- Ред {i} ---")
            pprint(rec, sort_dicts=False)

        if args.out:
            pd.DataFrame(records).to_excel(args.out, index=False)
            print(f"\n💾 Записано в: {args.out}")

    except Exception as e:
        print("❌ Грешка:", e)
        traceback.print_exc()
        sys.exit(1)
