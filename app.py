import re, unicodedata, logging
import pandas as pd
import pdfplumber
import streamlit as st
from io import BytesIO

logging.getLogger("pdfminer").setLevel(logging.ERROR)

# ---------- shared patterns ----------
ARABIC_RE = re.compile(r"[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]+")
ZERO_WIDTH_RE = re.compile(r"[\u200B-\u200D\u2060\uFEFF]")
ILLEGAL_XLSX_RE = re.compile(r"[\x00-\x08\x0B-\x0C\x0E-\x1F]")
META_PAT = (
    r"(?:^|\b)(Establishment Name|Establishment Number|Address|Category|"
    r"Scan QR|Printing Date|Total\s*:|QR\b|"
    r"المجموع|اسم المنشأة|رقم المنشأة|العنوان|الفئة|"
    r"امسح رمز الاستجابة السريعة|تاريخ الطباعة|إجمالي)(?:\b|:)"
)

def clean_cell(val):
    if val is None or (isinstance(val, float) and pd.isna(val)): return val
    s = str(val)
    s = unicodedata.normalize("NFKC", s)
    s = ZERO_WIDTH_RE.sub("", s)
    s = ARABIC_RE.sub("", s)          # strip Arabic content from cells (per your requirement)
    s = ILLEGAL_XLSX_RE.sub("", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def normalize_row(row): return [clean_cell(c) for c in row]

def extract_rows(file_like):
    rows_all = []
    with pdfplumber.open(file_like) as pdf:
        for page in pdf.pages:
            for t in (page.extract_tables() or []):
                t_norm = [normalize_row(r) for r in t if r and any(str(c or "").strip() for c in r)]
                rows_all.extend(t_norm)
    return rows_all

def drop_meta_headers(df, is_header_row_fn):
    def is_meta_row(row):
        joined = " | ".join(str(x) for x in row if pd.notna(x))
        return bool(re.search(META_PAT, joined, flags=re.IGNORECASE))
    df = df[~df.apply(lambda r: is_meta_row(list(r.values)), axis=1)]
    df = df[df.apply(lambda r: not is_header_row_fn(list(r.values)), axis=1)]
    return df

# ---------- Emirati cleaner (your original schema) ----------
EM_EXPECTED = [
    "Row No","Person Code","Person Name","Card Number",
    "Card Issue Date","Card Expiry Date","Job Type","Sex","Total Salary",
]
EM_EXPECTED_LC = {c.lower(): c for c in EM_EXPECTED}

def em_is_header_row(row):
    if not row or len(row) < 2: return False
    r0 = (str(row[0] or "")).lower().replace(" ", "")
    r1 = (str(row[1] or "")).lower().replace(" ", "")
    return ("rowno" in r0) and ("personcode" in r1)

def em_coerce_header(header):
    if not header: return EM_EXPECTED
    def k(x): return (x or "").lower().strip().replace("  ", " ")
    aliases = {
        "row no":"Row No","row":"Row No",
        "person code":"Person Code","person id":"Person Code",
        "person name":"Person Name","name":"Person Name",
        "card number":"Card Number",
        "card issue date":"Card Issue Date","issue date":"Card Issue Date",
        "card expiry date":"Card Expiry Date","expiry date":"Card Expiry Date",
        "job type":"Job Type","job":"Job Type",
        "sex":"Sex","gender":"Sex",
        "total salary":"Total Salary","salary":"Total Salary",
    }
    out=[]
    for c in header:
        cc=clean_cell(c); base=k(cc)
        out.append(EM_EXPECTED_LC.get(base, aliases.get(base, cc or "")))
    if sum(1 for c in out if c in EM_EXPECTED) < 5 and len(out) == len(EM_EXPECTED):
        return EM_EXPECTED
    return out

def to_clean_dataframe_emirati(file_like):
    rows = extract_rows(file_like)
    if not rows: raise ValueError("No tables found. If the PDF is scanned, please OCR first.")
    header, body = None, []
    for r in rows:
        if em_is_header_row(r):
            if header is None: header = r
            continue
        if any(str(c or "").strip() for c in r): body.append(r)
    header = em_coerce_header(header or [])
    body = [(r + [""]*max(0, len(header)-len(r)))[:len(header)] for r in body]
    df = pd.DataFrame(body, columns=header)

    df = drop_meta_headers(df, em_is_header_row)

    keep = [c for c in EM_EXPECTED if c in df.columns]
    if keep: df = df[keep]

    try: df = df.map(clean_cell)
    except AttributeError: df = df.applymap(clean_cell)

    # validations
    if "Row No" in df.columns:
        df = df[df["Row No"].astype(str).str.strip().str.isdigit()]
    if "Person Code" in df.columns:
        pc = df["Person Code"].astype(str).str.strip()
        df = df[pc.str.isdigit() & (pc.str.len() >= 10)]
    if "Sex" in df.columns:
        df["Sex"] = df["Sex"].astype(str).str.strip().str.title()
        df = df[df["Sex"].isin(["Male","Female"])]

    # drop empties & tidy columns
    df.dropna(how="all", inplace=True)
    df = df[~(df.astype(str).apply(lambda s: s.str.strip()).eq("").all(axis=1))]
    df.columns = [clean_cell(c) for c in df.columns]

    # types
    if "Row No" in df.columns:
        df["Row No"] = (df["Row No"].astype(str).str.strip().replace("", pd.NA).astype("Int64"))
    if "Person Code" in df.columns:
        df["Person Code"] = df["Person Code"].astype(str).str.strip()
    if "Total Salary" in df.columns:
        df["Total Salary"] = (
            df["Total Salary"].astype(str).str.replace(",", "", regex=False).str.strip()
            .replace("", pd.NA).astype(float)
        )
    # for col in ["Card Issue Date","Card Expiry Date"]:
    #     if col in df.columns:
    #         df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)
    return df

# ---------- Non-Emirati cleaner (new schema) ----------
NE_EXPECTED = [
    "Passport Number","Person Name","Card Type",
    "Job Name","Nationality","Card Number","Contract Type",
]
NE_EXPECTED_LC = {c.lower(): c for c in NE_EXPECTED}

def ne_is_header_row(row):
    if not row or len(row) < 2: return False
    r0 = (str(row[0] or "")).lower().replace(" ", "")
    r1 = (str(row[1] or "")).lower().replace(" ", "")
    return ("passport" in r0) and ("personname" in r1)

def ne_coerce_header(header):
    if not header: return NE_EXPECTED
    def k(x): return (x or "").lower().strip().replace("  ", " ")
    def strip_bilingual_noise(s):
        s = re.sub(r"\s*/\s*[\u0600-\u06FF].*$", "", s, flags=re.DOTALL)
        s = re.sub(r"\n+[\u0600-\u06FF].*$", "", s, flags=re.DOTALL)
        return s.strip()
    aliases = {
        "passport number":"Passport Number","passport no":"Passport Number",
        "رقم جواز السفر":"Passport Number","رقمجوازالسفر":"Passport Number",
        "person name":"Person Name","name":"Person Name","اسم الشخص":"Person Name",
        "card type":"Card Type","نوع البطاقة":"Card Type",
        "job name":"Job Name","job":"Job Name","المهنة":"Job Name",
        "nationality":"Nationality","الجنسية":"Nationality",
        "card number":"Card Number","رقم البطاقة":"Card Number",
        "contract type":"Contract Type","نوع العقد":"Contract Type",
    }
    out=[]
    for c in header:
        cc = clean_cell(c)
        cc = strip_bilingual_noise(cc)
        base = k(cc)
        out.append(NE_EXPECTED_LC.get(base, aliases.get(base, cc or "")))
    if sum(1 for c in out if c in NE_EXPECTED) < 5 and len(out) == len(NE_EXPECTED):
        return NE_EXPECTED
    return out

def to_clean_dataframe_non_emirati(file_like):
    rows = extract_rows(file_like)
    if not rows: raise ValueError("No tables found. If the PDF is scanned, please OCR first.")
    header, body = None, []
    for r in rows:
        if ne_is_header_row(r):
            if header is None: header = r
            continue
        if any(str(c or "").strip() for c in r): body.append(r)
    header = ne_coerce_header(header or [])
    body = [(r + [""]*max(0, len(header)-len(r)))[:len(header)] for r in body]
    df = pd.DataFrame(body, columns=header)

    df = drop_meta_headers(df, ne_is_header_row)

    keep = [c for c in NE_EXPECTED if c in df.columns]
    if keep: df = df[keep]

    try: df = df.map(clean_cell)
    except AttributeError: df = df.applymap(clean_cell)

    # --- NEW: split trailing ID from Person Name into Person Number ---
    if "Person Name" in df.columns:
        # capture: name (non-greedy) + trailing 8+ digits (if present) at the end of the cell
        split = df["Person Name"].str.extract(r'^(?P<name>.*?)(?P<number>\d{8,})\s*$', expand=True)
        # Person Number (text), Person Name cleaned
        df["Person Number"] = split["number"].fillna("")
        df["Person Name"] = split["name"].fillna(df["Person Name"]).str.strip()

    # place Person Number after Person Name if present
    if "Person Number" in df.columns:
        cols = list(df.columns)
        if "Person Name" in cols:
            cols.remove("Person Number")
            insert_at = cols.index("Person Name") + 1
            cols.insert(insert_at, "Person Number")
            df = df[cols]

    # validations/normalizers
    if "Passport Number" in df.columns:
        df["Passport Number"] = df["Passport Number"].astype(str).str.strip()
        df = df[df["Passport Number"] != ""]
    if "Contract Type" in df.columns:
        df["Contract Type"] = df["Contract Type"].astype(str).str.strip().str.title()
        common = {"Limited","Unlimited"}
        df = df[(df["Contract Type"] == "") | (df["Contract Type"].isin(common))]
    if "Card Number" in df.columns:
        df["Card Number"] = (
            df["Card Number"].astype(str).str.strip()
            .str.extract(r"(\d{5,})", expand=False).fillna("")  # keep first 5+ digit run
        )

    df.dropna(how="all", inplace=True)
    df = df[~(df.astype(str).apply(lambda s: s.str.strip()).eq("").all(axis=1))]
    df.columns = [clean_cell(c) for c in df.columns]

    for col in ["Passport Number","Card Number"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    return df

# ---------- UI ----------
st.set_page_config(page_title="MOHRE Employee Lists → Clean Excel", page_icon="📄")
st.title("MOHRE Employee Lists → Clean Excel")
st.caption("Upload a MOHRE PDF. The app removes Arabic text, QR codes, photos, and tidies the table into Excel format.")

tabs = st.tabs(["Emirati (Local Employee List)", "All Employees List"])

with tabs[0]:
    st.subheader("Emirati List")
    up_em = st.file_uploader("Upload MOHRE PDF (Emirati)", type=["pdf"], key="emirati_pdf")
    if up_em:
        with st.spinner("Cleaning Emirati list…"):
            try:
                df_em = to_clean_dataframe_emirati(up_em)
                st.success(f"Done. Rows: {len(df_em)}")
                st.dataframe(df_em.head(50), use_container_width=True)
                bio = BytesIO()
                with pd.ExcelWriter(bio, engine="openpyxl") as w:
                    df_em.to_excel(w, index=False, sheet_name="Employee_List")
                bio.seek(0)
                st.download_button(
                    "⬇️ Download Emirati Clean Excel",
                    data=bio.getvalue(),
                    file_name="Local_Employee_List_Cleaned.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(f"Conversion failed: {e}")

with tabs[1]:
    st.subheader("Non-Emirati List")
    up_ne = st.file_uploader("Upload MOHRE PDF (Non-Emirati)", type=["pdf"], key="non_emirati_pdf")
    if up_ne:
        with st.spinner("Cleaning Non-Emirati list…"):
            try:
                df_ne = to_clean_dataframe_non_emirati(up_ne)
                st.success(f"Done. Rows: {len(df_ne)}")
                st.dataframe(df_ne.head(50), use_container_width=True)
                bio = BytesIO()
                with pd.ExcelWriter(bio, engine="openpyxl") as w:
                    df_ne.to_excel(w, index=False, sheet_name="Employees_List")
                bio.seek(0)
                st.download_button(
                    "⬇️ Download Non-Emirati Clean Excel",
                    data=bio.getvalue(),
                    file_name="Employees_List_Cleaned.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(f"Conversion failed: {e}")
