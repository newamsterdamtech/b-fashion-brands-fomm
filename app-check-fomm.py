uploaded_files = st.file_uploader(
    "Upload multiple Excel files", type=["xlsx", "xls"], accept_multiple_files=True
)

def parse_percentage(val):
    if pd.isnull(val): return None
    if isinstance(val, (float, int)): return float(val)
    s = str(val).replace('%', '').replace(',', '.').strip()
    try: return float(s) / 100 if '%' in str(val) else float(s)
    except Exception: return None

def format_percentage(val):
    if pd.isnull(val): return ""
    try:
        percentage = float(val) * 100
        return f"{percentage:.2f}".replace('.', ',') + '%'
    except Exception: return ""

def clean_ean(val):
    if pd.isnull(val): return ""
    s = str(val)
    if s.endswith(".0"): s = s[:-2]
    return s.strip()

def get_header_mapping(columns):
    mapping = {}
    colmap = {re.sub(r'\W+', '', str(col)).upper(): col for col in columns}
    possible_cols = {
        "PO": ["PO"],
        "EAN CODES": ["EAN CODES", "Ean Codes", "EAN"],
        "PACKED": ["PACKED"],
        "ORDERED": ["ORDERED"],
        "PERCENTAGE": ["PERCENTAGE", "RATIO"],
    }
    for canon, variants in possible_cols.items():
        for var in variants:
            key = re.sub(r'\W+', '', var).upper()
            if key in colmap:
                mapping[canon] = colmap[key]
                break
    return mapping

def preprocess_excel_file_ean_codes(file):
    try:
        df = pd.read_excel(file, sheet_name='EAN Codes')
        df.dropna(how='all', inplace=True)
        df.columns = df.columns.str.strip()
        return {"EAN Codes": df}
    except Exception as e:
        return {}

def improved_preprocess_excel_file(file):
    xls = pd.ExcelFile(file)
    processed_sheets = {}
    for sheet_name in xls.sheet_names:
        df_raw = xls.parse(sheet_name, header=None)
        df_raw.dropna(how='all', inplace=True)
        possible_header_rows = df_raw.apply(
            lambda row: row.astype(str).str.contains('EAN', case=False, na=False).any(), axis=1
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
        xls = pd.ExcelFile(uploaded_file)
        if 'EAN Codes' in xls.sheet_names:
            sheets_data = preprocess_excel_file_ean_codes(uploaded_file)
        else:
            sheets_data = improved_preprocess_excel_file(uploaded_file)

        for df in sheets_data.values():
            header_map = get_header_mapping(df.columns)

            # Most flexible: Look for at least EAN and PACKED (with or without PO/ORDERED)
            ean_col = header_map.get("EAN CODES")
            packed_col = header_map.get("PACKED")
            po_col = header_map.get("PO")
            ordered_col = header_map.get("ORDERED")

            columns_to_use = []
            if po_col: columns_to_use.append(po_col)
            if ean_col: columns_to_use.append(ean_col)
            if ordered_col: columns_to_use.append(ordered_col)
            if packed_col: columns_to_use.append(packed_col)

            # Only process if we have at least EAN and PACKED
            if ean_col and packed_col:
                df_filtered = df[columns_to_use].copy()
                # Exclude rows with no valid PO, if PO exists; else just use EAN as identifier
                if po_col:
                    df_filtered = df_filtered[pd.to_numeric(df_filtered[po_col], errors='coerce').notnull()]
                    df_filtered[po_col] = df_filtered[po_col].astype(int)
                df_filtered[packed_col] = df_filtered[packed_col].fillna(0).astype(int)
                df_filtered = df_filtered[df_filtered[packed_col] != 0]
                df_filtered[ean_col] = df_filtered[ean_col].apply(clean_ean)
                # Rename for consistency
                col_rename = {}
                if po_col: col_rename[po_col] = "PO"
                if ean_col: col_rename[ean_col] = "EAN CODES"
                if ordered_col: col_rename[ordered_col] = "ORDERED"
                if packed_col: col_rename[packed_col] = "PACKED"
                df_filtered = df_filtered.rename(columns=col_rename)
                final_data.append(df_filtered)

            # Deviations
            percent_col = header_map.get("PERCENTAGE")
            if percent_col and (po_col or ean_col):
                df_copy = df.copy()
                if po_col:
                    df_copy = df_copy[pd.to_numeric(df_copy[po_col], errors='coerce').notnull()]
                df_copy['__percent'] = df_copy[percent_col].map(parse_percentage)
                deviations = df_copy[df_copy['__percent'] <= -0.05]
                if not deviations.empty:
                    deviations[percent_col] = deviations['__percent'].map(format_percentage)
                    deviations_data.append(deviations.drop(columns='__percent', errors='ignore'))

    # Output section
    if final_data:
        final_output_data = pd.concat(final_data, ignore_index=True)
        st.subheader("Processed Data")
        st.write(final_output_data.head(50))
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
        st.warning("No valid data found in uploaded files.")

    if deviations_data:
        deviations_output = pd.concat(deviations_data, ignore_index=True)
        st.subheader("Rows with Deviations (PERCENTAGE < -5%)")
        st.write(deviations_output.head(50))
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
        st.info("No rows with PERCENTAGE/RATIO below -5% found.")
