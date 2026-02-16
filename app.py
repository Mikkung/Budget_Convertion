import re
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st


RENAME_MAP = {
    "รหัสบัญชีงบประมาณ": "Budget_Account",
    "งบประมาณ": "Budget",
    "PR/กันงบ": "PR/Reserved_Budget",
    "ตั้งหนี้/จ่าย": "Accured/Paid",
    "คงเหลือ": "Remaining_Balance",
    "ใช้ไป%": "Spent%",
}


def extract_year_from_value(v) -> str:
    """Best-effort: return '2025' etc or '' if not found."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    # datetime-like
    try:
        if hasattr(v, "year"):
            y = int(v.year)
            if 1900 <= y <= 2100:
                return str(y)
    except Exception:
        pass

    s = str(v)
    m = re.search(r"(19|20)\d{2}", s)
    return m.group(0) if m else ""


def extract_year(raw: pd.DataFrame) -> str:
    """Try the notebook logic first, then scan a small area."""
    # Notebook: year = raw.iloc[5,0]
    try:
        y = extract_year_from_value(raw.iloc[5, 0])
        if y:
            return y
    except Exception:
        pass

    # Scan top-left area (first 25 rows x 6 cols)
    rmax = min(25, raw.shape[0])
    cmax = min(6, raw.shape[1])
    for r in range(rmax):
        for c in range(cmax):
            y = extract_year_from_value(raw.iloc[r, c])
            if y:
                return y
    return ""


def read_excel_any(uploaded_file) -> pd.DataFrame:
    """Read .xls or .xlsx robustly."""
    name = uploaded_file.name.lower()
    if name.endswith(".xls") and not name.endswith(".xlsx"):
        return pd.read_excel(uploaded_file, header=None, engine="xlrd")
    # default to openpyxl for xlsx
    return pd.read_excel(uploaded_file, header=None, engine="openpyxl")


def convert_budget(uploaded_file, keep_suffix_in_budget_code: bool = True) -> pd.DataFrame:
    raw = read_excel_any(uploaded_file)
    year = extract_year(raw)

    # --- delete first 9 rows; then delete "row 11" (original) => index 1 after reset
    df = raw.drop(index=list(range(0, 9)), errors="ignore").reset_index(drop=True)
    df = df.drop(index=[1], errors="ignore").reset_index(drop=True)

    # --- first remaining row becomes header
    df.columns = df.iloc[0].astype(str).str.strip()
    df = df.iloc[1:].reset_index(drop=True)

    # --- strip header names again (safer for thai headers with spaces)
    df.columns = pd.Index(df.columns).astype(str).str.strip()

    # --- add Year
    df["Year"] = year

    # --- rename columns
    df = df.rename(columns=RENAME_MAP)

    if "Budget_Account" not in df.columns:
        raise KeyError("Could not find 'รหัสบัญชีงบประมาณ' / Budget_Account column after header set.")

    # --- normalize text safely (keeps NA as NA)
    s = (df["Budget_Account"]
         .astype("string")
         .str.replace("\u00A0", " ", regex=False)  # non-breaking space
         .str.strip()
    )

    # --- group rows:
    # examples:
    #   G501 : something
    #   G501_1 : something
    #
    # If keep_suffix_in_budget_code=True => Budget_Code = G501_1
    # else => Budget_Code = G501
    if keep_suffix_in_budget_code:
        group_pat = r"^((?:G\d{3})(?:_[^:\s]+)?)\s*:\s*(.+)$"
    else:
        group_pat = r"^(G\d{3})(?:_[^:\s]+)?\s*:\s*(.+)$"

    m = s.str.extract(group_pat)  # col0=Budget_Code, col1=Budget_Type
    is_group = m[0].notna()

    df["Budget_Code"] = m[0].where(is_group).ffill()
    df["Budget_Type"] = m[1].where(is_group).ffill()

    # Remove group rows by blanking Budget_Account then dropping NA
    df["Budget_Account"] = s.where(~is_group)
    df = df.dropna(subset=["Budget_Account"]).reset_index(drop=True)

    # --- split item rows: "5xx1 ค่านู่น ..." => Expense_Code + Expense_Detail
    ba = df["Budget_Account"].astype("string").str.strip()
    df["Expense_Code"] = ba.str.split().str[0]
    df["Expense_Detail"] = ba.str.split().str[1:].str.join(" ")
    df = df.drop(columns=["Budget_Account"])

    # --- reorder columns (only keep those that exist)
    desired = [
        "Year", "Budget_Code", "Budget_Type",
        "Expense_Code", "Expense_Detail",
        "Budget", "PR/Reserved_Budget", "Accured/Paid", "Remaining_Balance", "Spent%"
    ]
    df = df[[c for c in desired if c in df.columns]]

    return df


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="converted")
    return bio.getvalue()


st.set_page_config(page_title="Budget File Converter", layout="wide")
st.title("Budget File Converter")
st.caption("Upload an Excel file (.xls / .xlsx) and download the converted file.")

uploaded = st.file_uploader("Upload file", type=["xls", "xlsx"])

keep_suffix = st.toggle("Keep suffix in Budget_Code (e.g., G501_1)", value=True)

if uploaded is not None:
    try:
        out_df = convert_budget(uploaded, keep_suffix_in_budget_code=keep_suffix)

        st.subheader("Preview")
        st.dataframe(out_df, use_container_width=True, height=420)

        out_name = Path(uploaded.name).stem + "_converted.xlsx"
        st.download_button(
            label="Download converted Excel",
            data=to_excel_bytes(out_df),
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Conversion failed: {e}")
        st.exception(e)
