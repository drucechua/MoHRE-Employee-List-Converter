import re, unicodedata, logging
import pandas as pd
import pdfplumber
import streamlit as st
from io import BytesIO

# Silence pdfminer noise
logging.getLogger("pdfminer").setLevel(logging.ERROR)

# --- Patterns ---
ARABIC_RE = re.compile(r"[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]+")
ZERO_WIDTH_RE = re.compile(r"[\u200B-\u200D\u2060\uFEFF]")
ILLEGAL_XLSX_RE = re.compile(r"[\x00-\x08\x0B-\x0C\x0E-\x1F]")

EXPECTED_COLS = [
    "Row No", "Person Code", "Person Name", "Card Number",
    "Card Issue Date", "Card Expiry Date", "Job Type", "Sex", "Total Salary",
]
EXPECTED_COLS_LC_MAP = {c.lower(): c for c in EXPECTED_COLS}

META_PAT = (
    r"(?:^|\b)(Establishment Name|Establishment Number|Address|Category|"
    r"Scan QR|Printing Date|Total\s*:|المجموع|QR\b)(?:\b|:)"
)

def clean_cell(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return val
    s = str(val)
    s = unicodedata.normalize("NFKC", s)
    s = ZERO_WIDTH_RE.sub("", s)
    s = ARABIC_RE.sub("", s)
    s = ILLEGAL_XLSX_RE.sub("", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def normalize_row(row): return [clean_cell(c) for c in row]

def is_header_row(row):
    if not row or len(row) < 2: return False
    r0 = (str(row[0] or "")).lower().replace(" ", "")
    r1 = (str(row[1] or "")).lower().replace(" ", "")
    return ("rowno" in r0) and ("personcode" in r1)

def coerce_header(header):
    if not header: return EXPECTED_COLS
    def k(x): return (x or "").lower().strip().replace("  ", " ")
    aliases = {
        "row no":"Row No","row":"Row No","person code":"Person Code","person id":"Person Code",
        "person name":"Person Name","name":"Person Name","card number":"Card Number",
        "card issue date":"Card Issue Date","issue date":"Card Issue Date",
        "card expiry date":"Card Expiry Date","expiry date":"Card Expiry Date",
        "job type":"Job Type","job":"Job Type","sex":"Sex","gender":"Sex",
        "total salary":"Total Salary","salary":"Total Salary",
    }
    out=[]
    for c in header:
        cc=clean_cell(c); base=k(cc)
        out.append(EXPECTED_COLS_LC_MAP.get(base, aliases.get(base, cc or "")))
    if sum(1 for c in out if c in EXPECTED_COLS)<5 and len(out)==len(EXPECTED_COLS):
        return EXPECTED_COLS
    return out

def extract_rows(file_like):
    rows_all=[]
    with pdfplumber.open(file_like) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables() or []
            for t in tables:
                t_norm=[normalize_row(r) for r in t if r and any(str(c or "").strip() for c in r)]
                rows_all.extend(t_norm)
    return rows_all

def to_clean_dataframe(file_like):
    rows = extract_rows(file_like)
    if not rows:
        raise ValueError("No tables found. If the PDF is scanned, please OCR first.")
    header, body = None, []
    for r in rows:
        if is_header_row(r):
            if header is None: header = r
            continue
        if any(str(c or "").strip() for c in r): body.append(r)
    header = coerce_header(header or [])
    body = [ (r + [""]*max(0,len(header)-len(r)))[:len(header)] for r in body ]
    df = pd.DataFrame(body, columns=header)

    # Drop meta rows
    def is_meta_row(row):
        joined = " | ".join(str(x) for x in row if pd.notna(x))
        return bool(re.search(META_PAT, joined, flags=re.IGNORECASE))
    df = df[~df.apply(lambda r: is_meta_row(list(r.values)), axis=1)]

    # Column order & basic validity
    keep = [c for c in EXPECTED_COLS if c in df.columns]
    if keep: df = df[keep]

    # Drop header-like body rows
    df = df[df.apply(lambda r: not is_header_row(list(r.values)), axis=1)]

    # Clean all cells
    try: df = df.map(clean_cell)
    except AttributeError: df = df.applymap(clean_cell)

    # Light validations
    if "Row No" in df.columns: df = df[df["Row No"].astype(str).str.strip().str.isdigit()]
    if "Person Code" in df.columns:
        df = df[df["Person Code"].astype(str).str.strip().str.isdigit() & (df["Person Code"].astype(str).str.len()>=10)]
    if "Sex" in df.columns:
        df["Sex"] = df["Sex"].astype(str).str.strip().str.title()
        df = df[df["Sex"].isin(["Male","Female"])]

    # Drop empty rows, sanitize columns
    df.dropna(how="all", inplace=True)
    df = df[~(df.astype(str).apply(lambda s: s.str.strip()).eq("").all(axis=1))]
    df.columns = [clean_cell(c) for c in df.columns]
    return df

st.set_page_config(page_title="MOHRE PDF → Clean Excel", page_icon="📄")
st.title("MOHRE Local Employee List → Clean Excel")
st.caption("Upload the MOHRE PDF. This app removes Arabic text, page headers, and images/QRs, then outputs a tidy Excel file.")

uploaded = st.file_uploader("Upload MOHRE PDF", type=["pdf"])
if uploaded:
    with st.spinner("Converting and cleaning…"):
        try:
            df = to_clean_dataframe(uploaded)
            st.success(f"Done. Rows: {len(df)}")

            st.dataframe(df.head(50), use_container_width=True)

            # Prepare download
            bio = BytesIO()
            with pd.ExcelWriter(bio, engine="openpyxl") as w:
                df.to_excel(w, index=False, sheet_name="Employee_List")
            bio.seek(0)
            st.download_button(
                label="⬇️ Download Clean Excel",
                data=bio.getvalue(),
                file_name="Local_Employee_List_Cleaned.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"Conversion failed: {e}")
else:
    st.info("Pick the original MOHRE PDF to begin.")
