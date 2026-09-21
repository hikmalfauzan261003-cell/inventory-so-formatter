import streamlit as st
import pandas as pd
import io

st.set_page_config(
    page_title="Inventory Data Formatter & SO Converter",
    page_icon="📦",
    layout="wide"
)

st.title("📦 Inventory Report to Template SO Formatter")
st.write("Upload file laporan inventaris mentah (`REPORT INVENTORY...`) untuk dikonversi & dirapikan sesuai format `TEMPLATE SO`.")

# Sidebar - Petunjuk & Pengaturan
st.sidebar.header("⚙️ Pengaturan & Petunjuk")
st.sidebar.info(
    """
    **Langkah Pemrosesan:**
    1. Upload file Excel/CSV data inventaris mentah.
    2. Sistem akan memetakan kolom secara otomatis (P/N, S/N, Location, Bin, Batch, Qty, dll).
    3. Pilih default Auditor dan Status jika tidak ada di file mentah.
    4. Unduh file hasil konversi yang sudah rapi.
    """
)

auditor_name = st.sidebar.text_input("Nama Auditor (Default)", value="AUDITOR")
so_date = st.sidebar.date_input("Tanggal SO", value=pd.to_datetime("today"))

# Upload File
uploaded_file = st.file_uploader(
    "Upload Data Utama (Excel atau CSV)", 
    type=["xlsx", "xls", "csv"]
)

if uploaded_file is not None:
    try:
        # Membaca Data
        if uploaded_file.name.endswith(".csv"):
            df_raw = pd.read_csv(uploaded_file, sep=None, engine="python")
        else:
            df_raw = pd.read_excel(uploaded_file)

        st.subheader("📋 Preview Data Mentah")
        st.dataframe(df_raw.head(10), use_container_width=True)

        # Standarisasi Nama Kolom Data Mentah (Capital & Strip)
        df_raw.columns = [str(c).strip().upper() for c in df_raw.columns]

        # Kamus Pemetaan Kolom Mentah -> Kolom Template SO
        mapping_rules = {
            'LOC': ['LOC', 'LOCATION', 'STORE', 'WAREHOUSE'],
            'BIN': ['BIN', 'LOCATION BIN', 'BINNING'],
            'BATCH': ['BATCH', 'GRB', 'BATCH NO', 'BATCH_NO'],
            'PN': ['PN', 'PART NO', 'PART NUMBER', 'P/N'],
            'SN': ['SN', 'SERIAL NO', 'SERIAL NUMBER', 'S/N'],
            'PN DESCRIPTION': ['PN DESCRIPTION', 'DESCRIPTION', 'PART DESCRIPTION', 'ITEM DESCRIPTION'],
            'QTY EMRO': ['QTY AVAILABLE', 'QTY EMRO', 'QTY', 'QUANTITY', 'QTY_AVAILABLE'],
            'SHELF LIFE EXP': ['SHELF LIFE EXPIRATION', 'SHELF LIFE EXP', 'EXPIRY DATE', 'EXP DATE'],
            'UOM': ['UOM', 'UNIT', 'UNIT OF MEASURE'],
            'CATEGORY GROUP': ['CATEGORY', 'CATEGORY GROUP', 'TYPE'],
            'CONDITION': ['CONDITION', 'COND', 'PART CONDITION']
        }

        # Proses Transformasi Data
        df_transformed = pd.DataFrame()

        # 1. Peta / Salin kolom yang tersedia
        for target_col, possible_sources in mapping_rules.items():
            matched_col = next((col for col in possible_sources if col in df_raw.columns), None)
            if matched_col:
                df_transformed[target_col] = df_raw[matched_col]
            else:
                df_transformed[target_col] = ""

        # 2. Tambahkan & Kalkulasi Kolom Wajib Template SO
        df_transformed['COUNT NO'] = range(1, len(df_transformed) + 1)
        df_transformed['QTY ACTUAL'] = df_transformed['QTY EMRO']  # Default disamakan sebelum di-audit
        
        # Konversi numerik untuk kalkulasi selisih (Diff)
        qty_emro_num = pd.to_numeric(df_transformed['QTY EMRO'], errors='coerce').fillna(0)
        qty_act_num = pd.to_numeric(df_transformed['QTY ACTUAL'], errors='coerce').fillna(0)
        
        df_transformed['DIFF'] = qty_act_num - qty_emro_num
        df_transformed['RESULT'] = df_transformed['DIFF'].apply(lambda x: 'MATCHED' if x == 0 else ('MINUS' if x < 0 else 'SURPLUS'))
        df_transformed['STATUS'] = df_transformed['RESULT'].apply(lambda x: 'MATCHED' if x == 'MATCHED' else 'OPEN')
        
        df_transformed['DATE'] = so_date.strftime("%d/%m/%Y")
        df_transformed['AUDITOR'] = auditor_name
        df_transformed['REMARK'] = ""
        df_transformed['PENYELESAIAN'] = ""
        df_transformed['CORRECTIVE ACTION'] = ""
        df_transformed['REASON'] = ""

        # Susun Ulang Kolom Sesuai Urutan Standard Template SO
        final_columns = [
            'COUNT NO', 'LOC', 'BIN', 'BATCH', 'PN', 'SN', 
            'PN DESCRIPTION', 'QTY EMRO', 'QTY ACTUAL', 'DIFF', 
            'RESULT', 'STATUS', 'SHELF LIFE EXP', 'UOM', 'DATE', 
            'AUDITOR', 'REMARK', 'PENYELESAIAN', 'CATEGORY GROUP', 
            'CORRECTIVE ACTION', 'REASON'
        ]

        df_final = df_transformed[final_columns]

        st.success("✅ Data berhasil ditransformasi dan disesuaikan dengan Template SO!")

        st.subheader("📊 Preview Data Hasil Transformasi")
        st.dataframe(df_final.head(10), use_container_width=True)

        # Export ke File Excel menggunakan BytesIO
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name='SO_Rapi')
            
            # Styling otomatis lebar kolom
            worksheet = writer.sheets['SO_Rapi']
            for col in worksheet.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = col[0].column_letter
                worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

        output_buffer.seek(0)

        # Tombol Download
        st.download_button(
            label="📥 Download Excel Template SO (Rapi)",
            data=output_buffer,
            file_name=f"RESULT_STOCK_OPNAME_{so_date.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {e}")