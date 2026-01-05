# streamlit_app.py
"""
SettleMate - Streamlit app: UPI settlement summarizer
Supports:
- .xlsx
- real legacy .xls
- .xls files that are actually .xlsx (wrong extension)
"""

import streamlit as st
import pandas as pd
import io, zipfile, re
from datetime import datetime

# ---------------- PAGE CONFIG ----------------
st.set_page_config(page_title="SettleMate", layout="wide")

# ---------------- LOGIN ----------------
def check_password():
    def password_entered():
        if st.session_state["password"] == st.secrets["password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input("Enter password:", type="password", on_change=password_entered, key="password")
        st.stop()
    elif not st.session_state["password_correct"]:
        st.text_input("Enter password:", type="password", on_change=password_entered, key="password")
        st.error("😕 Password incorrect")
        st.stop()

check_password()

# ---------------- UI HEADER ----------------
st.title("SettleMate — UPI Settlement Summarizer")
st.caption("Upload a ZIP or Excel files (.xls / .xlsx). Files are auto-detected safely.")

# ---------------- UTILITIES ----------------
date_regex = re.compile(
    r'(\d{4}-\d{2}-\d{2})|(\d{2}-\d{2}-\d{4})|([A-Za-z]{3,9}\s+\d{1,2},\s*\d{4})'
)

def is_xlsx_bytes(bts: bytes) -> bool:
    return bts[:4] == b'PK\x03\x04'

def read_excel_bytes_safe(bts, fname):
    try:
        if is_xlsx_bytes(bts):
            return pd.read_excel(io.BytesIO(bts), sheet_name=None, engine="openpyxl", header=None)
        else:
            return pd.read_excel(io.BytesIO(bts), sheet_name=None, engine="xlrd", header=None)
    except Exception as e:
        raise RuntimeError(f"Failed reading Excel file {fname}: {e}")

# ---------------- ZIP EXTRACTION ----------------
def extract_excel_files_from_zip(zip_bytes):
    excel_files = []
    z = zipfile.ZipFile(io.BytesIO(zip_bytes))

    for name in z.namelist():
        base = name.split('/')[-1]
        if not base or base.startswith(('._', '~$')):
            continue
        if base.lower().endswith(('.xls', '.xlsx', '.xlsm', '.xltx', '.xltm')):
            excel_files.append((base, z.read(name)))
    return excel_files

# ---------------- HELPERS ----------------
def find_header_row(raw_df):
    nrows = min(10, len(raw_df))
    for i in range(nrows):
        row = raw_df.iloc[i].astype(str).str.lower().tolist()
        if any(('description' in c) or ('no of txns' in c) or ('debit' in c) or
               ('credit' in c) or ('transaction' in c) or ('total fee' in c) or
               ('total transaction' in c) for c in row):
            return i
    return 0

def extract_date_from_raw(raw):
    for col in raw.columns:
        for v in raw[col].astype(str).head(10):
            m = date_regex.search(v)
            if m:
                return m.group(0)
    return ''

def normalize_cols(cols):
    return [str(c).strip().lower() for c in cols]

def safe_numeric(series):
    return pd.to_numeric(
        series.astype(str).str.replace(',', '').str.replace('₹', '').str.strip(),
        errors='coerce'
    ).fillna(0.0)

# ---------------- SHEET TOTALS ----------------
def extract_sheet_level_totals(raw, header_idx):
    res = {
        'sheet_fee_debit': 0.0,
        'sheet_fee_credit': 0.0,
        'sheet_transaction_debit': 0.0,
        'sheet_transaction_credit': 0.0
    }

    if header_idx >= len(raw):
        return res

    header = raw.iloc[header_idx].astype(str).str.lower().tolist()
    data_row_idx = header_idx + 1
    if data_row_idx >= len(raw):
        return res

    row = raw.iloc[data_row_idx].tolist()

    def get_two_numbers(start):
        nums = []
        for v in row[start:start + 6]:
            try:
                s = str(v).replace(',', '').replace('₹', '')
                nums.append(float(re.findall(r'\d+\.?\d*', s)[0]))
                if len(nums) == 2:
                    break
            except:
                continue
        return nums

    for i, h in enumerate(header):
        if 'total fee' in h:
            nums = get_two_numbers(i)
            if nums:
                res['sheet_fee_debit'] += nums[0]
                if len(nums) > 1:
                    res['sheet_fee_credit'] += nums[1]
            break

    for i, h in enumerate(header):
        if 'total transaction' in h:
            nums = get_two_numbers(i)
            if nums:
                res['sheet_transaction_debit'] += nums[0]
                if len(nums) > 1:
                    res['sheet_transaction_credit'] += nums[1]
            break

    return res

# ---------------- MAIN PROCESSING ----------------
def process_excel_bytes(file_bytes, filename, fee_keywords, txn_keywords, debug=False):
    totals = {
        'file': filename,
        'date': '',
        'total_fee_debit': 0.0,
        'total_fee_credit': 0.0,
        'total_transaction_debit': 0.0,
        'total_transaction_credit': 0.0,
        'sheet_fee_debit': 0.0,
        'sheet_fee_credit': 0.0,
        'sheet_transaction_debit': 0.0,
        'sheet_transaction_credit': 0.0
    }

    sheets = read_excel_bytes_safe(file_bytes, filename)

    for sheet_name, raw in sheets.items():
        if raw is None or raw.empty:
            continue

        header_idx = find_header_row(raw)
        headers = raw.iloc[header_idx].fillna('').astype(str)
        df = raw.iloc[header_idx + 1:].copy().reset_index(drop=True)
        df.columns = normalize_cols(headers)

        if not totals['date']:
            totals['date'] = extract_date_from_raw(raw)

        sheet_totals = extract_sheet_level_totals(raw, header_idx)
        for k in sheet_totals:
            totals[k] += sheet_totals[k]

        desc_col = next((c for c in df.columns if 'description' in c), df.columns[0])
        debit_col = next((c for c in df.columns if 'debit' in c), None)
        credit_col = next((c for c in df.columns if 'credit' in c), None)

        if not debit_col or not credit_col:
            continue

        df['debit'] = safe_numeric(df[debit_col])
        df['credit'] = safe_numeric(df[credit_col])
        desc = df[desc_col].astype(str).str.lower()

        fee_pattern = '|'.join(map(re.escape, fee_keywords)) or 'fee'
        txn_pattern = '|'.join(map(re.escape, txn_keywords)) or 'transaction amount'

        totals['total_fee_debit'] += df.loc[desc.str.contains(fee_pattern), 'debit'].sum()
        totals['total_fee_credit'] += df.loc[desc.str.contains(fee_pattern), 'credit'].sum()
        totals['total_transaction_debit'] += df.loc[desc.str.contains(txn_pattern), 'debit'].sum()
        totals['total_transaction_credit'] += df.loc[desc.str.contains(txn_pattern), 'credit'].sum()

    return totals

# ---------------- UI ----------------
st.sidebar.header("Options")
fee_kw = [x.strip().lower() for x in st.sidebar.text_input(
    "Fee keywords (comma separated)", "fee, switching fee, psp fee"
).split(",") if x.strip()]

txn_kw = [x.strip().lower() for x in st.sidebar.text_input(
    "Transaction keywords (comma separated)", "transaction amount, approved transaction amount"
).split(",") if x.strip()]

uploaded = st.file_uploader(
    "Upload ZIP or Excel files",
    type=["zip", "xls", "xlsx"],
    accept_multiple_files=True
)

if uploaded and st.button("Process files"):
    files = []
    for u in uploaded:
        data = u.read()
        if u.name.lower().endswith('.zip'):
            files.extend(extract_excel_files_from_zip(data))
        else:
            files.append((u.name, data))

    results = []
    for fname, bts in files:
        try:
            results.append(process_excel_bytes(bts, fname, fee_kw, txn_kw))
        except Exception as e:
            st.error(f"{fname}: {e}")

    if results:
        out_df = pd.DataFrame(results)

        st.subheader("Per-file summary")
        st.dataframe(out_df, use_container_width=True)

        overall = {
            'file': 'ALL_FILES_SUM',
            'date': out_df['date'].mode().iloc[0] if not out_df['date'].empty else '',
            'total_fee_debit': round(out_df['total_fee_debit'].sum(), 2),
            'total_fee_credit': round(out_df['total_fee_credit'].sum(), 2),
            'total_transaction_debit': round(out_df['total_transaction_debit'].sum(), 2),
            'total_transaction_credit': round(out_df['total_transaction_credit'].sum(), 2),
            'sheet_fee_debit': round(out_df['sheet_fee_debit'].sum(), 2),
            'sheet_fee_credit': round(out_df['sheet_fee_credit'].sum(), 2),
            'sheet_transaction_debit': round(out_df['sheet_transaction_debit'].sum(), 2),
            'sheet_transaction_credit': round(out_df['sheet_transaction_credit'].sum(), 2),
        }

        st.subheader("Overall totals (all uploaded files)")
        st.table(pd.DataFrame([overall]).T.rename(columns={0: 'value'}))

        csv = out_df.to_csv(index=False, float_format='%.2f').encode('utf-8')
        st.download_button("Download summary CSV", csv, "settlement_summary.csv", "text/csv")
