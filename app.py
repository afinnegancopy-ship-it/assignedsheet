import streamlit as st
import pandas as pd
import re
import difflib
from datetime import datetime
import unicodedata
from io import BytesIO

# -------- CONFIG --------
special_num_map = {
    "50 R":"50 R",
    "39":"39 (15.5in)",
    "8":"UK 8",
    "6":"UK 6",
    "16":"16in",
    "32 B":"",
    "31":"W 31",
    "27":"W 27",
    "32/32":"W 32 L 32",
    "32 32":"W 32 L 32",
    "26 32":"W 26 L 32",
    "26/32":"W 26 L 32",
    "34/30":"W 34 L 30",
    "34 30":"W 34 L 30",
    "6 Months":"",
    "6yrs":"",
    "27 32":"W 27 L 32",
    "Med/Lge":"M/L",
    "155":"15.5in",
    "ONESIZE":"",
    "ONE SIZE":"",
    "41":"41 (15.75in)",
    "One Size":"",
    "one size":"",
    "NOSIZ":"",
    "Sml/Med":"S/M",
    "80": "",
}

# --- Helper functions ---
def find_col(df, target_name):
    target_norm = re.sub(r"\W+", "", target_name).lower()
    for c in df.columns:
        if re.sub(r"\W+", "", str(c)).lower() == target_norm:
            return c
    return None

def parse_excel_date(val):
    if pd.isna(val) or val == "":
        return None
    if isinstance(val, (datetime, pd.Timestamp)):
        return val
    if isinstance(val, (int, float)):
        try:
            return pd.to_datetime('1899-12-30') + pd.to_timedelta(val, unit='D')
        except:
            return None
    for fmt in ("%d-%m-%Y", "%d.%m.%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(str(val), fmt)
        except:
            continue
    return None

def normalize_name(name):
    if pd.isna(name) or not name: return ""
    name = str(name).lower()
    name = re.sub(r"[-_]", " ", name)
    name = re.sub(r"\s+", " ", name).strip()
    misspell_map = {"eamon": "eamonn", "emma rose": "emma-rose"}
    return misspell_map.get(name, name)

def normalize_brand(name):
    if pd.isna(name) or not name: return ""
    name = str(name).lower()
    name = ''.join(c for c in unicodedata.normalize('NFKD', name) if not unicodedata.combining(c))
    name = re.sub(r"[-_]", " ", name)
    name = re.sub(r"\s+", " ", name).strip()
    words = name.split()
    words.sort()
    return " ".join(words)

def sort_by_store_and_description(df, description_col):
    if 'Date+Store' in df.columns and description_col in df.columns:
        # --- Safe store extraction ---
        def extract_store(val):
            try:
                val = str(val)
                parts = val.split(' ', 1)
                return parts[1] if len(parts) > 1 else ''
            except:
                return ''
        df['_store'] = df['Date+Store'].apply(extract_store)
        # Sort Z→A by store, A→Z by description
        df.sort_values(by=['_store', description_col], ascending=[False, True], inplace=True)
        df.drop(columns=['_store'], inplace=True)
    return df

# --- Streamlit UI ---
st.title("Studio Pull")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
selected_date = st.date_input("Select the date to filter from")

if uploaded_file:
    st.info("Reading Excel file...")
    xls = pd.ExcelFile(uploaded_file)
    
    st.write("Sheets detected:", xls.sheet_names)
    
    # --- Load Brand Size Mapping ---
    if "Brand Size Mapping" in xls.sheet_names:
        brand_size_map_df = pd.read_excel(xls, sheet_name="Brand Size Mapping", dtype=str)
        st.write("Brand Size Mapping loaded. Sample:", brand_size_map_df.head())
    else:
        st.error("Sheet 'Brand Size Mapping' not found in Excel.")
        st.stop()

    brand_size_map = {}
    for _, row in brand_size_map_df.iterrows():
        brand = str(row[0]).strip().lower()
        raw_size = str(row[1]).strip()
        conv_size = str(row[2]).strip()
        if brand not in brand_size_map:
            brand_size_map[brand] = {}
        brand_size_map[brand][raw_size] = conv_size
    
    # --- Load main sheet ---
    search_tab_name = "Search_Report_Download"
    designers_sheet_name = "Designers"
    
    if search_tab_name not in xls.sheet_names:
        st.error(f"Sheet '{search_tab_name}' not found in Excel.")
        st.stop()
    
    df = pd.read_excel(xls, sheet_name=search_tab_name, dtype=object)
    st.write("Main sheet loaded. Sample data:", df.head())
    
    # --- Detect columns ---
    photo_model_col = find_col(df, "Photo Model Date")
    photo_still_col = find_col(df, "Photo Still Date")
    dept_col = find_col(df, "Department")
    item_store_col = find_col(df, "Item Store Flag")
    ppid_col = find_col(df, "PPID")
    description_col = find_col(df, "Description") or find_col(df, "Retek Description")
    barcode_col = find_col(df, "Barcode")
    model_col = find_col(df, "Model Name")
    brand_col = find_col(df, "Brand")
    size_val_col = find_col(df, "Size")

    # --- Ensure Size and Description are strings ---
    if size_val_col in df.columns:
        df[size_val_col] = df[size_val_col].astype(str).replace('nan', '')
    if description_col in df.columns:
        df[description_col] = df[description_col].astype(str)

    # --- Filter by date ---
    assigned_dt = pd.Timestamp(selected_date)
    if photo_model_col:
        df = df[df[photo_model_col].apply(lambda x: (parse_excel_date(x) is None) or (parse_excel_date(x) >= assigned_dt))]
    if photo_still_col:
        df = df[df[photo_still_col].apply(lambda x: (parse_excel_date(x) is None) or (parse_excel_date(x) >= assigned_dt))]
    
    if df.empty:
        st.warning("No rows remaining after date filter.")
    
    # --- Load Model Information ---
    if 'Model Information' in xls.sheet_names:
        model_info_df = pd.read_excel(xls, sheet_name='Model Information', dtype=object)
        model_info_name_col = model_info_df.columns[0]
        model_info_conv_col = model_info_df.columns[1]
        model_conv_map = {str(k).lower(): str(v) for k, v in zip(model_info_df[model_info_name_col], model_info_df[model_info_conv_col])}
    else:
        model_conv_map = {}

    # --- Prioritize rows with Model Name ---
    df['has_model'] = df[model_col].notna() & (df[model_col].astype(str).str.strip() != "")
    df.sort_values(by=['has_model'], ascending=False, inplace=True)
    df = df.drop_duplicates(subset=[ppid_col], keep='first')
    df.drop(columns=['has_model'], inplace=True)
    
    # --- Conversion + Size ---
    def generate_conversion_size(row):
        model_name = str(row[model_col]).strip() if pd.notna(row.get(model_col)) else ""
        size = str(row[size_val_col]).strip() if pd.notna(row.get(size_val_col)) else ""
        brand = str(row[brand_col]).strip() if pd.notna(row.get(brand_col)) else ""
        
        normalized_model = normalize_name(model_name)
        conv_model = model_conv_map.get(normalized_model, "")
        if not conv_model:
            best_match = difflib.get_close_matches(normalized_model, model_conv_map.keys(), n=1, cutoff=0.8)
            if best_match:
                conv_model = model_conv_map[best_match[0]]
        
        normalized_brand = normalize_brand(brand)
        normalized_brand_map = {normalize_brand(k): k for k in brand_size_map.keys()}
        brand_matches = difflib.get_close_matches(normalized_brand, normalized_brand_map.keys(), n=1, cutoff=0.7)
        if brand_matches:
            actual_brand_key = normalized_brand_map[brand_matches[0]]
            if size in brand_size_map[actual_brand_key]:
                size = brand_size_map[actual_brand_key][size]
        
        final_size = special_num_map.get(size, size)
        return f"{conv_model} {final_size}".strip() if final_size else ""
    
    df['Conversion + Size'] = df.apply(generate_conversion_size, axis=1)
    
    # --- Split Designers vs Original ---
    designers_mask = df[dept_col].astype(str).str.strip().eq("Designers") if dept_col else pd.Series(False, index=df.index)
    designers_df = df[designers_mask].copy()
    orig_df_mod = df[~designers_mask].copy()
    
    # --- Add Date+Store ---
    for d_df in [designers_df, orig_df_mod]:
        store_col = item_store_col if item_store_col else None
        d_df['Date+Store'] = assigned_dt.strftime("%d.%m.%Y") + " " + d_df[store_col].astype(str) if store_col else assigned_dt.strftime("%d.%m.%Y")
    
    # --- Prepare output ---
    designers_out_cols = ['Date+Store', ppid_col, description_col, 'Conversion + Size']
    if barcode_col: designers_out_cols.insert(3, barcode_col)
    if size_val_col: designers_out_cols.insert(4, size_val_col)
    designers_df = designers_df[[c for c in designers_out_cols if c in designers_df.columns]]
    
    orig_out_cols = ['Date+Store', ppid_col, description_col, size_val_col, 'Conversion + Size']
    orig_df_mod = orig_df_mod[[c for c in orig_out_cols if c in orig_df_mod.columns]]
    orig_df_mod.insert(3, 'Blank', '')

    designers_df.dropna(inplace=True)
    orig_df_mod.dropna(inplace=True)
    
    # --- Sort by store Z-A, then Description A-Z safely ---
    designers_df = sort_by_store_and_description(designers_df, description_col)
    orig_df_mod = sort_by_store_and_description(orig_df_mod, description_col)
    
    # --- Save to BytesIO for download ---
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        orig_df_mod.to_excel(writer, sheet_name=search_tab_name, index=False)
        designers_df.to_excel(writer, sheet_name=designers_sheet_name, index=False)
    output.seek(0)
    
    st.success("Processing complete! Download below:")
    st.download_button(
        label="Download Processed Excel",
        data=output,
        file_name=f"Processed_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
