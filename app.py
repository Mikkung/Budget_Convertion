import io
import pandas as pd
import streamlit as st

from converter import convert_budget_file

st.set_page_config(page_title="Budget File Converter", layout="wide")
st.title("Budget File Converter")
st.caption("Upload an Excel file (.xls/.xlsx) → convert → preview → download. (SharePoint upload disabled)")

uploaded = st.file_uploader("Upload your budget Excel file", type=["xlsx", "xls"])

sheet_index = st.number_input("Sheet index to convert (0 = first sheet)", min_value=0, value=0, step=1)

if uploaded is not None:
    with st.spinner("Converting..."):
        try:
            df_converted, meta = convert_budget_file(uploaded, sheet_name=int(sheet_index))
        except Exception as e:
            st.error(f"Conversion failed: {e}")
            st.stop()

    year = meta.get("year") or ""
    st.success("Converted successfully.")

    c1, c2, c3 = st.columns([1,1,2])
    with c1:
        st.metric("Rows", len(df_converted))
    with c2:
        st.metric("Columns", df_converted.shape[1])
    with c3:
        if year:
            st.write(f"Detected Year: **{year}**")

    st.subheader("Preview")
    st.dataframe(df_converted, use_container_width=True, height=420)

    st.subheader("Download")
    col1, col2 = st.columns(2)

    with col1:
        out_xlsx = io.BytesIO()
        with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
            df_converted.to_excel(writer, index=False, sheet_name="Converted")
        st.download_button(
            "Download Excel (.xlsx)",
            data=out_xlsx.getvalue(),
            file_name=f"converted_{year}.xlsx" if year else "converted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with col2:
        out_csv = df_converted.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "Download CSV (.csv)",
            data=out_csv,
            file_name=f"converted_{year}.csv" if year else "converted.csv",
            mime="text/csv",
        )
else:
    st.info("Upload a file to begin.")
