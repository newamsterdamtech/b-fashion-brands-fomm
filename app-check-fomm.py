import streamlit as st
import pandas as pd
from io import BytesIO
import re

st.title("B Fashion Brands - Fomm Excel Processing")

uploaded_files = st.file_uploader(
    "Upload multiple Excel files", type=["xlsx"], accept_multiple_files=True
)

def parse_percentage(val):
    """Converts string percentage to float (e.g. '-14,29%' -> -0.1429). Handles comma and dot decimals."""
    if pd.isnull(val):
        return None
    if isinstance(val, (float, int)):
        return float(val)
    s = str(val).replace('%', '').replace(',', '.').strip()
    try:
        return float(s) / 100 if '%' in str(val) else float(s)
    except Exception:
        return None

def format_percentage(val):
    """Formats a float like -0.142857 to '-14,29%' with comma as decimal and two decimals."""
    if pd.isnull(val):
        return ""
    try:
        percentage = float(val) * 100
        # Use round and format, then replace decimal dot with comma
        formatted = f"{percentage:.2f}".replace('.', ',') + '%'
        return formatted
    except Exception:
        return ""

def improved_preprocess_excel_file(file):
    xls = pd.ExcelFile(file)
    processed_sheets = {}
    for sheet_name in xls.sheet_names:
        df_raw = xls.parse(sheet_name, header=None)
        df_raw.dropna(how='all', inplace=True)
        possible_header_rows = df_raw.apply(
            lambda row: row.astype(str).str.contains('EAN CODES', case=False, na=False).any(), axis=1
        )
        if possible_header_rows.any():
            header_row_idx = possible_header_rows.idxmax()
            df = pd.read_excel(file, sheet_name=sheet_name, header=header_row_idx)
            df.dropna(how='all', inplace=True)
            df.columns = df.columns.str.strip()
            processed_sheets[sheet_name] = df
    return processed_sheets

if uploaded_files:
    final_data = []
    deviations_data = []
    for uploaded_file in uploaded_files:
        sheets_data = improved_preprocess_excel_file(uploaded_file)
        for df in sheets_data.values():
            # First output: Only PO, EAN CODES, PACKED
            if {'PO', 'EAN CODES', 'PACKED'}.issubset(df.columns):
                df_filtered = df[['PO', 'EAN CODES', 'PACKED']].copy()
                df_filtered = df_filtered[pd.to_numeric(df_filtered['PO'], errors='coerce').notnull()]
                df_filtered = df_filtered[pd.to_numeric(df_filtered['PACKED'], errors='coerce').notnull()]
                df_filtered['PO'] = df_filtered['PO'].astype(int)
                df_filtered['PACKED'] = df_filtered['PACKED'].astype(int)
                df_filtered = df_filtered[df_filtered['PACKED'] != 0]
                final_data.append(df_filtered)

            # Second output: Any row where 'PERCENTAGE' < -0.05
            # Find column name, case-insensitive match for 'PERCENTAGE'
            percent_col = next((col for col in df.columns if re.sub(r'\s+', '', col).upper() == 'PERCENTAGE'), None)
            if percent_col:
                df_copy = df.copy()
                df_copy['__percent'] = df_copy[percent_col].map(parse_percentage)
                deviations = df_copy[df_copy['__percent'] <= -0.05]  # -5%
                if not deviations.empty:
                    # Format the percentage column for display/export
                    deviations[percent_col] = deviations['__percent'].map(format_percentage)
                    deviations_data.append(deviations.drop(columns='__percent', errors='ignore'))

    # Show PO/EAN CODES/PACKED table preview and download
    if final_data:
        final_output_data = pd.concat(final_data, ignore_index=True)
        st.subheader("Processed Data (PO, EAN CODES, PACKED)")
        st.write(final_output_data.head())
        output = BytesIO()
        final_output_data.to_csv(output, index=False)
        output.seek(0)
        st.download_button(
            "Download Output CSV",
            data=output,
            file_name="Inputfile.csv",
            mime="text/csv"
        )
    else:
        st.warning("No valid PO/EAN CODES/PACKED data found in uploaded files.")

    # Show deviation rows preview
    if deviations_data:
        deviations_output = pd.concat(deviations_data, ignore_index=True)
        st.subheader("Rows with Deviations (PERCENTAGE < -5%)")
        st.write(deviations_output.head(50))  # Show up to 50 rows for preview
        deviations_csv = BytesIO()
        deviations_output.to_csv(deviations_csv, index=False)
        deviations_csv.seek(0)
        st.download_button(
            "Download Deviations CSV",
            data=deviations_csv,
            file_name="Afwijkende-percentages.csv",
            mime="text/csv"
        )
    else:
        st.info("No rows with PERCENTAGE below -5% found.")
