#!/usr/bin/env python3
"""
Convert a MOHRE 'Local Employee List' PDF to a clean Excel file.

Actions:
1) Extract tabular data from every page (ignores images/QRs automatically).
2) Strip Arabic text from headers and cells (keep English, numbers, symbols).
3) Remove repeated headers that appear on each page.
4) Write a single, tidy Excel sheet.

Usage:
  python clean_employee_pdf.py "local Employee List 03-10-2025.pdf" -o "Local_Employee_List_Cleaned.xlsx"
"""

import argparse
import logging
import re
import unicodedata
from pathlib import Path

import pandas as pd
import pdfplumber

# Silence noisy font warnings from pdfminer/pdfplumber
logging.getLogger("pdfminer").setLevel(logging.ERROR)

# --- Patterns & constants -----------------------------------------------------

# One regex set (no duplicates)
ARABIC_RE = re.compile(r"[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]+")
ZERO_WIDTH_RE = re.compile(r"[\u200B-\u200D\u2060\uFEFF]")
ILLEGAL_XLSX_RE = re.compile(r"[\x00-\x08\x0B-\x0C\x0E-\x1F]")  # Excel-forbidden control chars

EXPECTED_COLS = [
    "Row No",
    "Person Code",
    "Person Name",
    "Card Number",
    "Card Issue Date",
    "Card Expiry Date",
    "Job Type",
    "Sex",
    "Total Salary",
]

EXPECTED_COLS_LC_MAP = {c.lower(): c for c in EXPECTED_COLS}

# --- Cleaners -----------------------------------------------------------------

def clean_cell(val):
    """
    Normalize Unicode; remove Arabic, zero-width, and Excel-illegal chars.
    Keep English, digits, punctuation; collapse whitespace.
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return val
    s = str(val)
    s = unicodedata.normalize("NFKC", s)
    s = ZERO_WIDTH_RE.sub("", s)
    s = ARABIC_RE.sub("", s)
    s = ILLEGAL_XLSX_RE.sub("", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def normalize_row(row):
    """Apply clean_cell to every cell in a row list."""
    return [clean_cell(c) for c in row]

# --- Header detection/mapping -------------------------------------------------

def is_header_row(row):
    """
    Heuristic: a header row contains 'Row No' and 'Person Code' (after cleaning).
    Null-safe and tolerant of spacing/casing.
    """
    if not row or len(row) < 2:
        return False
    r0 = (str(row[0] or "")).lower().replace(" ", "")
    r1 = (str(row[1] or "")).lower().replace(" ", "")
    return ("rowno" in r0) and ("personcode" in r1)

def coerce_header(header):
    """
    Map the extracted header to the expected column names (robust to minor OCR glitches).
    """
    if not header:
        return EXPECTED_COLS

    def keyize(x: str) -> str:
        return (x or "").lower().strip().replace("  ", " ")

    aliases = {
        "row no": "Row No",
        "row": "Row No",
        "person code": "Person Code",
        "person id": "Person Code",
        "person name": "Person Name",
        "name": "Person Name",
        "card number": "Card Number",
        "card issue date": "Card Issue Date",
        "issue date": "Card Issue Date",
        "card expiry date": "Card Expiry Date",
        "expiry date": "Card Expiry Date",
        "job type": "Job Type",
        "job": "Job Type",
        "sex": "Sex",
        "gender": "Sex",
        "total salary": "Total Salary",
        "salary": "Total Salary",
    }

    mapped = []
    for col in header:
        col_clean = clean_cell(col)
        base = keyize(col_clean)
        if base in EXPECTED_COLS_LC_MAP:
            out = EXPECTED_COLS_LC_MAP[base]
        else:
            out = aliases.get(base, col_clean or "")
        mapped.append(out)

    # If mapping confidence is low but widths match, force expected header.
    if sum(1 for c in mapped if c in EXPECTED_COLS) < 5 and len(mapped) == len(EXPECTED_COLS):
        return EXPECTED_COLS

    return mapped

# --- Extraction ---------------------------------------------------------------

def extract_tables(pdf_path: Path):
    """Extract tables from each page using pdfplumber's table extraction."""
    rows_all = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()
            if not tables:
                continue
            for t in tables:
                # Clean each row immediately
                t_norm = [normalize_row(r) for r in t if r and any((c is not None and str(c).strip()) for c in r)]
                rows_all.extend(t_norm)
    return rows_all

def split_header_and_body(rows):
    """
    Find the first 'real' header and drop any subsequent header repeats.
    Return (header, body_rows).
    """
    header = None
    body = []
    for row in rows:
        if is_header_row(row):
            if header is None:
                header = row
            continue  # skip repeats
        # keep only non-empty rows
        if any((str(cell or "").strip()) for cell in row):
            body.append(row)
    return header, body

def align_to_header(body, header_len):
    """Ensure each body row has exactly header_len columns (pad or trim)."""
    fixed = []
    for r in body:
        if len(r) < header_len:
            r = r + [""] * (header_len - len(r))
        elif len(r) > header_len:
            r = r[:header_len]
        fixed.append(r)
    return fixed

# --- Main ---------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf", type=str, help="Path to the PDF file")
    ap.add_argument("-o", "--output", type=str, default=None, help="Output Excel path (.xlsx)")
    args = ap.parse_args()

    pdf_path = Path(args.pdf)
    if not pdf_path.exists():
        raise SystemExit(f"❌ File not found: {pdf_path}")

    out_path = Path(args.output) if args.output else pdf_path.with_name(pdf_path.stem + "_Cleaned.xlsx")

    # 1) Extract raw rows from all pages (images ignored)
    rows = extract_tables(pdf_path)
    if not rows:
        raise SystemExit("❌ No tables found. If this PDF is scanned, run OCR first (e.g., ocrmypdf).")

    # 2) Separate header from body, remove header repetition
    raw_header, body = split_header_and_body(rows)

    # 3) Build the final header (clean + map to expected)
    header = coerce_header(raw_header or [])
    header_len = len(header)

    # 4) Align body rows to header width
    body = align_to_header(body, header_len)

    # 5) Create DataFrame and final cleanups
    df = pd.DataFrame(body, columns=header)

    # Keep only the columns we care about, preserving order when present
    keep_cols = [c for c in EXPECTED_COLS if c in df.columns]
    if keep_cols:
        df = df[keep_cols]

    # Drop any header-like rows that slipped into body (paranoia)
    df = df[df.apply(lambda r: not is_header_row(list(r.values)), axis=1)]

    # Apply cleaner to every cell — pandas 2.2+: DataFrame.map; fallback to applymap
    try:
        df = df.map(clean_cell)
    except AttributeError:
        df = df.applymap(clean_cell)

    # Drop fully-empty rows after cleaning
    df.dropna(how="all", inplace=True)
    is_blank = df.astype(str).apply(lambda s: s.str.strip()).eq("").all(axis=1)
    df = df[~is_blank]

    # Final column sanitization
    df.columns = [clean_cell(c) for c in df.columns]

    # --- Drop MOHRE page metadata blocks that sometimes get extracted as rows ---
    META_PAT = (
        r"(?:^|\b)(Establishment Name|Establishment Number|Address|Category|"
        r"Scan QR|Printing Date|Total\s*:|المجموع|QR\b)(?:\b|:)"
    )

    def is_meta_row(row) -> bool:
        joined = " | ".join(str(x) for x in row if pd.notna(x))
        return bool(re.search(META_PAT, joined, flags=re.IGNORECASE))

    df = df[~df.apply(lambda r: is_meta_row(list(r.values)), axis=1)]

    # --- (Optional but helpful) sanity checks to keep only employee rows ---
    # 1) Row No should be integer-like
    def intlike(s): 
        s = str(s).strip()
        return s.isdigit()

    if "Row No" in df.columns:
        df = df[df["Row No"].apply(intlike)]

    # 2) Person Code is usually a long numeric string (>=10 digits)
    if "Person Code" in df.columns:
        df = df[df["Person Code"].apply(lambda x: str(x).strip().isdigit() and len(str(x).strip()) >= 10)]

    # 3) Sex should be Male/Female if present
    if "Sex" in df.columns:
        df["Sex"] = df["Sex"].astype(str).str.strip().str.title()
        df = df[df["Sex"].isin(["Male", "Female"])]


    # 6) Write Excel
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Employee_List", index=False)

    print(f"✅ Done: {out_path}")

if __name__ == "__main__":
    main()
