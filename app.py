import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.title("Upload & Transform Excel Downtime")
st.subheader("Semangat Ges! 🚀")

# Upload file
uploaded_file = st.file_uploader("Upload file Excel", type=["xlsx"])

def transform_excel(file):
    df = pd.read_excel(file, header=None)

    # HEADER (Koordinat tetap mengikuti struktur file lama Anda)
    tanggal_row = df.iloc[0, 5:].ffill()
    shift_row   = df.iloc[1, 5:].ffill()
    sku_row     = df.iloc[2, 5:]
    desc_row    = df.iloc[3, 5:]

    # LEVEL
    levels = df.iloc[4:, 1:5]
    levels.columns = ['L1', 'L2', 'L3', 'L4']
    levels = levels.ffill(axis=0)

    # DATA
    data_values = df.iloc[4:, 5:]

    final_list = []

    for col_idx in range(data_values.shape[1]):
        tgl = tanggal_row.iloc[col_idx]
        sft = shift_row.iloc[col_idx]
        sku = sku_row.iloc[col_idx]
        dsc = desc_row.iloc[col_idx]

        for row_idx in range(data_values.shape[0]):
            durasi = data_values.iloc[row_idx, col_idx]

            if pd.notna(durasi) and str(durasi).replace('.','',1).isdigit():
                if float(durasi) > 0:
                    final_list.append({
                        'Title': dsc,          # SKU asal jadi Title
                        'Item Code': sku,      # Deskripsi asal jadi Item Code
                        'Tanggal': tgl,
                        'Shift': sft,
                        'Level 1': levels.iloc[row_idx, 0],
                        'Level 2': levels.iloc[row_idx, 1],
                        'Level 3': levels.iloc[row_idx, 2],
                        'Level 4': levels.iloc[row_idx, 3],
                        'Durasi2': durasi      # Nama kolom durasi diubah sesuai request
                    })

    df_bersih = pd.DataFrame(final_list)
    df_bersih = df_bersih.drop_duplicates()

    # --- TAMBAH KOLOM KOSONG (REKAP SESUAI REQUEST) ---
    # Daftar kolom yang diinginkan (termasuk yang kosong)
    kolom_final = [
        "Title", "Item Code", "SKU2", "Item Code 2", "Tanggal", "Shift", 
        "Waktu", "Plant", "Line", "Batch Mix", "Output-Ca", "Output2-Ca", 
        "Output-Pcs", "Output2-Pcs", "Volume", "OEE", "Downtime", "Planned_D", 
        "Durasi", "Level 1", "Level 2", "Level 3", "Level 4", "Durasi2", 
        "Komen", "Act_CU", "Max_Cu", "Max Output", "Speed", "Uptime"
    ]

    # Buat kolom yang belum ada dengan isi kosong
    for col in kolom_final:
        if col not in df_bersih.columns:
            df_bersih[col] = np.nan

    # Urutkan kolom sesuai list kolom_final
    df_bersih = df_bersih[kolom_final]

    return df_bersih

if uploaded_file:
    st.success("File berhasil diupload!")
    df_hasil = transform_excel(uploaded_file)

    st.write("Preview hasil (30 Kolom):")
    st.dataframe(df_hasil.head())

    # --- PROSES SIMPAN SEBAGAI EXCEL TABLE (CTRL+T) ---
    output = BytesIO()
    # Gunakan xlsxwriter sebagai engine
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_hasil.to_excel(writer, index=False, sheet_name='Downtime_Table')
        
        workbook  = writer.book
        worksheet = writer.sheets['Downtime_Table']
        
        # Dapatkan dimensi tabel
        (max_row, max_col) = df_hasil.shape
        
        # Buat daftar kolom untuk format tabel
        column_settings = [{'header': column} for column in df_hasil.columns]
        
        # Tambahkan fungsi Table (Ctrl+T)
        worksheet.add_table(0, 0, max_row, max_col - 1, {
            'columns': column_settings,
            'style': 'TableStyleMedium9' # Warna biru standar Excel Table
        })
        
        # Auto-fit kolom biar rapi
        for i, col in enumerate(df_hasil.columns):
            worksheet.set_column(i, i, 15)

    output.seek(0)

    st.download_button(
        label="Download hasil Excel (Format Table)",
        data=output,
        file_name="hasil_transform_sic.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )