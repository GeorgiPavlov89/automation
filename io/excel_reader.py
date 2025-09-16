# automation/io/excel_reader.py
from __future__ import annotations
import re
from pathlib import Path
import pandas as pd

# Мап към нормализирани ключове
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
FLAGS = [v for v in MAP.values() if v.startswith("do_")]
TEXT_TRUE = {"1", "1.0", "да", "д", "yes", "y", "true", "t", "✓", "x"}

def is_valid_egn(s: str) -> bool:
    s = (s or "").strip()
    if not re.fullmatch(r"\d{10}", s):
        return False
    w = [2,4,8,5,10,9,7,3,6]
    chk = sum(int(s[i]) * w[i] for i in range(9)) % 11
    if chk == 10: chk = 0
    return chk == int(s[9])

def read_cases(path: str) -> list[dict]:
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Не намирам файла: {p}")

    # Избор на engine по разширение
    ext = p.suffix.lower()
    engine = "xlrd" if ext == ".xls" else ("openpyxl" if ext in (".xlsx", ".xlsm") else None)
    print(f"[excel_reader] Чета: {p} (engine={engine})")

    df = pd.read_excel(p, engine=engine)  # pandas.read_excel поддържа xlrd/openpyxl
    print(f"[excel_reader] Колони (източник): {list(df.columns)}")

    # Почисти заглавията и мапни
    df = df.rename(columns=lambda c: str(c).strip())
    df = df.rename(columns={c: MAP.get(c, c) for c in df.columns})
    print(f"[excel_reader] Колони (норм.): {list(df.columns)}")

    # Увери се, че всички флагове ги има
    for f in FLAGS:
        if f not in df.columns:
            df[f] = 0

    # Нормализирай флаговете към 0/1
    for f in FLAGS:
        df[f] = df[f].apply(lambda v: 1 if str(v).strip().lower() in TEXT_TRUE else v)
    df[FLAGS] = df[FLAGS].apply(pd.to_numeric, errors="coerce").fillna(0).astype(int)

    # Базова проверка за ЕГН
    bad_rows = []
    for i, v in df["egn_or_eik"].astype(str).items():
        t = v.strip()
        if t.isdigit() and len(t) == 10 and not is_valid_egn(t):
            bad_rows.append(i+2)
    if bad_rows:
        print("⚠ Невалидна ЕГН чексума в редове:", bad_rows)

    return df.fillna("").to_dict(orient="records")

if __name__ == "__main__":
    import argparse, pprint, sys, traceback
    ap = argparse.ArgumentParser(description="Прочети и нормализирай Excel с дела")
    ap.add_argument("path", help="Път до .xls/.xlsx")
    args = ap.parse_args()
    try:
        recs = read_cases(args.path)
        print(f"\n✅ Прочетени редове: {len(recs)}")
        for i, r in enumerate(recs, 1):
            print(f"\n--- Ред {i} ---")
            pprint.pprint(r, sort_dicts=False)
    except Exception as e:
        print("❌ Грешка:", e)
        traceback.print_exc()
        sys.exit(1)
