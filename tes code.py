import streamlit as st
import pandas as pd
import openpyxl
import io

st.set_page_config(
    page_title="Stock Opname Template Inserter",
    page_icon="📑",
    layout="wide"
)

st.title("📑 Stock Opname Auto-Formatter (Preserve Template & Sheets)")
st.write("Aplikasi ini menginjeksikan data inventaris ke sheet **Worksheet** mulai baris ke-8 pada `TEMPLATE SO_2.xlsx` tanpa mengganggu baris 1-7 serta sheet `NLA` & `UNRECORDED`.")

# 1. Component File Uploader
col1, col2 = st.columns(2)
with col1:
    template_file = st.file_uploader("1. Upload File Template (`TEMPLATE SO_2.xlsx`)", type=["xlsx"])
with col2:
    data_file = st.file_uploader("2. Upload Data Utama / Mentah (`REPORT INVENTORY...`)", type=["xlsx", "xls", "csv"])

if template_file is not None and data_file is not None:
    try:
        # Load Template Workbook
        wb = openpyxl.load_workbook(template_file)
        
        if "Worksheet" not in wb.sheetnames:
            st.error("Sheet 'Worksheet' tidak ditemukan pada file template!")
        else:
            ws = wb["Worksheet"]
            
            # Load Data Mentah
            if data_file.name.endswith(".csv"):
                df_raw = pd.read_csv(data_file, sep=None, engine="python")
            else:
                df_raw = pd.read_excel(data_file)
                
            st.subheader("📋 Preview Data Mentah")
            st.dataframe(df_raw.head(5), use_container_width=True)
            
            # Standarisasi Nama Kolom Data Mentah
            df_raw.columns = [str(c).strip().upper() for c in df_raw.columns]
            
            # 18 Kolom Target Sesuai Kebutuhan
            target_cols = [
                "LOCATION", "BIN", "GRB", "BATCH", "PN", "SN", 
                "PN DESCRIPTION", "QTY AVAILABLE", "QTY RESERVED", 
                "QTY IN TRANSFER", "QTY PENDING RI", "QTY US", 
                "QTY IN REPAIR", "SHELF LIFE EXPIRATION", 
                "TOOL LIFE EXPIRATION", "CONDITION", "CATEGORY", 
                "OWNER", "UOM"
            ]
            
            # Pemetaan Otomatis Variasi Nama Kolom Mentah -> Target Column
            mapping_rules = {
                "LOCATION": ["LOCATION", "LOC", "STORE", "WAREHOUSE"],
                "BIN": ["BIN", "LOCATION BIN", "BINNING"],
                "GRB": ["GRB", "GRB NO", "GRB_NO"],
                "BATCH": ["BATCH", "BATCH NO", "BATCH_NO"],
                "PN": ["PN", "PART NO", "PART NUMBER", "P/N"],
                "SN": ["SN", "SERIAL NO", "SERIAL NUMBER", "S/N"],
                "PN DESCRIPTION": ["PN DESCRIPTION", "DESCRIPTION", "PART DESCRIPTION", "ITEM DESCRIPTION"],
                "QTY AVAILABLE": ["QTY AVAILABLE", "QTY_AVAILABLE", "QTY", "QUANTITY", "QTY EMRO"],
                "QTY RESERVED": ["QTY RESERVED", "QTY_RESERVED"],
                "QTY IN TRANSFER": ["QTY IN TRANSFER", "QTY_IN_TRANSFER", "QTY IN TF"],
                "QTY PENDING RI": ["QTY PENDING RI", "QTY_PENDING_RI"],
                "QTY US": ["QTY US", "QTY_US"],
                "QTY IN REPAIR": ["QTY IN REPAIR", "QTY_IN_REPAIR"],
                "SHELF LIFE EXPIRATION": ["SHELF LIFE EXPIRATION", "SHELF LIFE EXP", "EXPIRY DATE"],
                "TOOL LIFE EXPIRATION": ["TOOL LIFE EXPIRATION", "TOOL LIFE EXP"],
                "CONDITION": ["CONDITION", "COND", "PART CONDITION"],
                "CATEGORY": ["CATEGORY", "CATEGORY GROUP", "TYPE"],
                "OWNER": ["OWNER", "OWNERSHIP"],
                "UOM": ["UOM", "UNIT", "UNIT OF MEASURE"]
            }
            
            # Buat DataFrame Terstruktur
            df_mapped = pd.DataFrame()
            for target in target_cols:
                possible_sources = mapping_rules.get(target, [target])
                matched_col = next((col for col in possible_sources if col in df_raw.columns), None)
                if matched_col:
                    df_mapped[target] = df_raw[matched_col]
                else:
                    df_mapped[target] = ""
                    
            # Hapus isi data lama di Worksheet mulai dari baris ke-8 kebawah
            max_row = ws.max_row
            if max_row >= 8:
                ws.delete_rows(8, max_row - 7 + 1)
                
            # Tulis data baru mulai baris ke-8
            start_row = 8
            for row_idx, row_data in enumerate(df_mapped.values, start=start_row):
                for col_idx, val in enumerate(row_data, start=1):
                    # Penanganan nilai NaN / None
                    cell_val = None if pd.isna(val) else val
                    ws.cell(row=row_idx, column=col_idx, value=cell_val)
                    
            # Export Ke Memory Buffer
            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)
            
            st.success("✅ Data berhasil diinjeksikan! Baris 1-7 serta Sheet NLA & UNRECORDED tetap utuh.")
            
            st.download_button(
                label="📥 Download Template SO Terisi (.xlsx)",
                data=output_buffer,
                file_name="TEMPLATE_SO_UPDATED.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {e}")