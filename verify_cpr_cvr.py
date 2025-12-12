#!/usr/bin/env python3
from pathlib import Path
import csv
import re
from datetime import date

FILENAME_REGEX = re.compile(r"^\d{4}-\d{2}-\d{2}_Fakturaer_\d{4}\.csv$")

# ---------------- CPR helpers ----------------

def _cpr_century(yy: int, d7: int) -> int | None:
    """
    Map (YY, 7th digit) -> century per CPR table.
    Returns 1800, 1900, or 2000; None if no mapping.
    Table (from CPR/Wikipedia):
      if d7 in [0,1,2,3]:        1900-1999
      if d7 == 4:                2000-2036 when YY 00-36 else 1900-1999 (37-99)
      if d7 in [5,6,7,8]:        2000-2057 when YY 00-57 else 1800-1899 (58-99)
      if d7 == 9:                2000-2036 when YY 00-36 else 1900-1999 (37-99)
    """
    if 0 <= d7 <= 3:
        return 1900
    if d7 == 4:
        return 2000 if 0 <= yy <= 36 else 1900
    if 5 <= d7 <= 8:
        return 2000 if 0 <= yy <= 57 else 1800
    if d7 == 9:
        return 2000 if 0 <= yy <= 36 else 1900
    return None

def _is_valid_date(d: int, m: int, y: int) -> bool:
    try:
        date(y, m, d)
        return True
    except ValueError:
        return False

def cpr_parse_and_checks(num: str):
    """
    Return dict with:
      plausible_by_date: bool
      birthdate: 'YYYY-MM-DD' or ''
      mod11_pass: bool
    """
    out = {"plausible_by_date": False, "birthdate": "", "mod11_pass": False}
    if not (len(num) == 10 and num.isdigit()):
        return out

    dd = int(num[0:2])
    mm = int(num[2:4])
    yy = int(num[4:6])
    d7 = int(num[6])

    century = _cpr_century(yy, d7)
    if century is None:
        return out
    yyyy = century + yy

    if not _is_valid_date(dd, mm, yyyy):
        return out

    out["plausible_by_date"] = True
    out["birthdate"] = f"{yyyy:04d}-{mm:02d}-{dd:02d}"

    # Historical Mod-11 check (not mandatory post-2007)
    weights = [4,3,2,7,6,5,4,3,2,1]
    total = sum(int(d) * w for d, w in zip(num, weights))
    out["mod11_pass"] = (total % 11 == 0)
    return out

# ---------------- CVR helper ----------------

def cvr_is_valid(num: str) -> bool:
    """
    CVR (8 digits) with Mod-11 check using weights [2,7,6,5,4,3,2,1].
    """
    if not (len(num) == 8 and num.isdigit()) or num == "00000000":
        return False
    weights = [2,7,6,5,4,3,2,1]
    total = sum(int(d) * w for d, w in zip(num, weights))
    return total % 11 == 0

# ---------------- CSV utilities ----------------

def read_first_row_second_col(csv_path: Path) -> str:
    """
    Return the value of the 2nd column from the first non-empty row.
    Tries to sniff delimiter; falls back to comma.
    """
    with csv_path.open("r", newline="") as f:
        sample = f.read(2048)
        f.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=[",",";","\t","|"])
        except csv.Error:
            dialect = csv.get_dialect("excel")
        reader = csv.reader(f, dialect)
        for row in reader:
            if not row or all(not c.strip() for c in row):
                continue
            return row[1] if len(row) > 1 else ""
    return ""

def clean_number(value: str) -> str:
    """
    Keep digits only, then remove ONE leading '0' if present.
    (Matches your original requirement.)
    """
    digits = re.sub(r"\D", "", value or "")
    return digits[2:] if digits.startswith("00") else digits

# ---------------- Main ----------------

def main():
    cwd = Path(__file__).resolve().parent
    inputs = sorted(p for p in cwd.iterdir() if p.is_file() and FILENAME_REGEX.match(p.name))

    results = []
    for csv_path in inputs:
        raw = read_first_row_second_col(csv_path)
        cleaned = clean_number(raw)

        cpr = cpr_parse_and_checks(cleaned)
        cvr_ok = cvr_is_valid(cleaned)

        # Derived “type” label for convenience
        types = []
        if cpr["plausible_by_date"]:
            types.append("CPR")
        if cvr_ok:
            types.append("CVR")
        type_label = " & ".join(types) if types else "neither"

        results.append({
            "filename": csv_path.name,
            "raw_second_column": raw,
            "cleaned_number": cleaned,
            "cpr_plausible_by_date": "yes" if cpr["plausible_by_date"] else "no",
            "cpr_birthdate": cpr["birthdate"],
            "cpr_mod11_pass": "yes" if cpr["mod11_pass"] else "no",
            "valid_cvr": "yes" if cvr_ok else "no",
            "classified_as": type_label,
        })

    out_path = cwd / "faktura_id_check_results.csv"
    with out_path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=[
                "filename",
                "raw_second_column",
                "cleaned_number",
                "cpr_plausible_by_date",
                "cpr_birthdate",
                "cpr_mod11_pass",
                "valid_cvr",
                "classified_as",
            ],
        )
        writer.writeheader()
        writer.writerows(results)

    print(f"Wrote {len(results)} rows to {out_path.name}")

if __name__ == "__main__":
    main()
