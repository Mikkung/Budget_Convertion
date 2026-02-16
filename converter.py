from __future__ import annotations

import io
import re
from typing import Dict, Tuple, Union, Optional

import pandas as pd

# Map Thai column headers -> standardized column names
THAI_TO_STD = {
    "รหัสบัญชีงบประมาณ": "Budget_Account",
    "งบประมาณ": "Budget",
    "PR/กันงบ": "PR/Reserved_Budget",
    "ตั้งหนี้/จ่าย": "Accured/Paid",
    "คงเหลือ": "Remaining_Balance",
    "ใช้ไป%": "Spent%",
}

STD_COLS = [
    "Year",
    "Budget_Code",
    "Budget_Type",
    "Expense_Code",
    "Expense_Detail",
    "Budget",
    "PR/Reserved_Budget",
    "Accured/Paid",
    "Remaining_Balance",
    "Spent%",
]

def _read_excel_any(file_like, sheet_name=0) -> pd.DataFrame:
    """
    Read both .xlsx and legacy .xls from an uploaded file.
    """
    # streamlit gives a BytesIO-like object; pandas can read it directly.
    try:
        return pd.read_excel(file_like, header=None, sheet_name=sheet_name, engine="openpyxl")
    except Exception:
        # fallback for .xls
        return pd.read_excel(file_like, header=None, sheet_name=sheet_name, engine="xlrd")

def _detect_year(df_raw: pd.DataFrame) -> Optional[str]:
    """
    Notebook logic used: year = df.iloc[5,0] then take 2nd token.
    We'll try that first, then fall back to scanning for a 4-digit year.
    """
    try:
        v = df_raw.iloc[5, 0]
        s = str(v)
        toks = s.split()
        if len(toks) >= 2:
            y = toks[1]
            if re.fullmatch(r"\d{4}", y):
                return y
        # sometimes the year might be directly in the cell
        m = re.search(r"\b(19|20)\d{2}\b", s)
        if m:
            return m.group(0)
    except Exception:
        pass

    # fallback scan top-left block
    block = df_raw.iloc[:20, :5].astype("string")
    flat = " ".join([x for x in block.to_numpy().ravel() if x and x != "nan"])
    m = re.search(r"\b(19|20)\d{2}\b", flat)
    return m.group(0) if m else None

def _find_header_row(df_raw: pd.DataFrame) -> int:
    """
    Prefer finding the row that contains the Thai header 'รหัสบัญชีงบประมาณ'.
    Fallback to the notebook's known structure (start at row 9).
    """
    for i in range(min(50, len(df_raw))):
        row = df_raw.iloc[i].astype("string").fillna("")
        if any("รหัสบัญชีงบประมาณ" in str(x) for x in row.tolist()):
            return i
    return 9  # notebook: drop rows 0-8 then set header

def _clean_numeric(series: pd.Series) -> pd.Series:
    """
    Convert numbers stored as strings like '1,234.00' to numeric.
    Leaves non-convertible as NA.
    """
    s = series.astype("string")
    s = s.str.replace(",", "", regex=False).str.replace("฿", "", regex=False).str.strip()
    return pd.to_numeric(s, errors="coerce")

def convert_budget_file(
    file_like: Union[io.BytesIO, bytes],
    sheet_name=0,
) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """
    Convert the raw budget excel into a standardized table:

    Year | Budget_Code | Budget_Type | Expense_Code | Expense_Detail | Budget | PR/Reserved_Budget
         | Accured/Paid | Remaining_Balance | Spent%

    Replicates your notebook logic, but with a safer header/year detection.

    Returns (df_converted, meta)
    """
    df_raw = _read_excel_any(file_like, sheet_name=sheet_name)
    year = _detect_year(df_raw)

    # ---- Trim to header region ----
    header_row = _find_header_row(df_raw)
    df = df_raw.iloc[header_row:].copy().reset_index(drop=True)

    # Notebook: drop row index 1 after trimming (often a blank or metadata row)
    if len(df) > 2:
        df = df.drop(index=1).reset_index(drop=True)

    # Make first row header
    df.columns = df.iloc[0]
    df = df.iloc[1:].reset_index(drop=True)

    # Drop fully empty rows
    df = df.dropna(how="all").reset_index(drop=True)

    # Add Year
    df["Year"] = year if year is not None else pd.NA

    # Rename Thai columns -> standard
    df = df.rename(columns=THAI_TO_STD)

    # Ensure required base columns exist
    for col in ["Budget_Account", "Budget", "PR/Reserved_Budget", "Accured/Paid", "Remaining_Balance", "Spent%"]:
        if col not in df.columns:
            df[col] = pd.NA

    # ---- Group parsing (same pattern as notebook) ----
    s = (
        df["Budget_Account"]
        .astype("string")
        .str.replace("\u00A0", " ", regex=False)  # NBSP -> space
        .str.strip()
    )

    # group: Gxxx or Gxxx_suffix : type
    g_pat = r"^((?:G\d{3})(?:_[^:\s]+)?)\s*:\s*(.+)$"
    m = s.str.extract(g_pat)  # m[0]=Budget_Code, m[1]=Budget_Type
    is_g_group = m[0].notna()

    # text-only group: non-empty, not starting with digit, and not G-group
    starts_with_digit = s.str.match(r"^\d", na=False)
    is_text_group = s.notna() & (s != "") & (~starts_with_digit) & (~is_g_group)

    # Budget_Code: from G-group only, then forward fill
    df["Budget_Code"] = m[0].where(is_g_group).ffill()

    # Budget_Type: from G-group (after ':') OR text group (whole line), then forward fill
    df["Budget_Type"] = pd.Series(pd.NA, index=df.index, dtype="string")
    df.loc[is_g_group, "Budget_Type"] = m.loc[is_g_group, 1]
    df.loc[is_text_group, "Budget_Type"] = s.loc[is_text_group]
    df["Budget_Type"] = df["Budget_Type"].ffill()

    # Remove group rows from detail lines
    is_group = is_g_group | is_text_group
    df["Budget_Account"] = s.where(~is_group)
    df = df.dropna(subset=["Budget_Account"]).reset_index(drop=True)

    # ---- Split Expense_Code / Expense_Detail ----
    df["Expense_Code"] = df["Budget_Account"].astype("string").str.split().str[0]
    df["Expense_Detail"] = df["Budget_Account"].astype("string").str.split().str[1:].str.join(" ")
    df = df.drop(columns=["Budget_Account"])

    # ---- Clean numeric columns (optional but recommended) ----
    for c in ["Budget", "PR/Reserved_Budget", "Accured/Paid", "Remaining_Balance"]:
        df[c] = _clean_numeric(df[c])

    # Keep Spent% as-is (often already numeric/percent); attempt numeric conversion gently
    if "Spent%" in df.columns:
        sp = df["Spent%"].astype("string").str.replace("%", "", regex=False).str.strip()
        df["Spent%"] = pd.to_numeric(sp, errors="ignore")

    # Ensure all output columns
    for c in STD_COLS:
        if c not in df.columns:
            df[c] = pd.NA
    df = df[STD_COLS]

    meta = {"year": year or ""}
    return df, meta
