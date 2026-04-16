import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.title("Upload & Transform Excel Downtime")
st.subheader("Semangat Ges!")

# Upload file
uploaded_file = st.file_uploader("Upload file Excel", type=["xlsx"])

def transform_excel(file):
    df = pd.read_excel(file, header=None)

    # HEADER
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
                        'SKU': sku,
                        'Deskripsi': dsc,
                        'Tanggal': tgl,
                        'Shift': sft,
                        'Level 1': levels.iloc[row_idx, 0],
                        'Level 2': levels.iloc[row_idx, 1],
                        'Level 3': levels.iloc[row_idx, 2],
                        'Level 4': levels.iloc[row_idx, 3],
                        'Durasi': durasi
                    })

    df_bersih = pd.DataFrame(final_list)
    df_bersih = df_bersih.drop_duplicates()

    return df_bersih

# Kalau file sudah diupload
if uploaded_file:
    st.success("File berhasil diupload!")

    df_hasil = transform_excel(uploaded_file)

    st.write("Preview hasil:")
    st.dataframe(df_hasil)

    # Convert ke Excel (in-memory)
    output = BytesIO()
    df_hasil.to_excel(output, index=False)
    output.seek(0)

    # Tombol download
    st.download_button(
        label="Download hasil Excel",
        data=output,
        file_name="hasil_transform.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )