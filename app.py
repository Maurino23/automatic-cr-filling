import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="Automated CR Filling", layout="wide")

# Tambahkan logo/banner
st.image("Logo App.png", use_column_width=True)

st.title("📋 Automated CR Filling App")
st.caption("Versi UI interaktif dengan preview, validasi, dan unduhan hasil otomatis")

# Tombol refresh halaman
if st.button("🔄 Refresh Aplikasi"):
    for key in st.session_state.keys():
        del st.session_state[key]
    try:
        st.rerun()
    except:
        try:
            st.experimental_rerun()
        except:
            st.warning("🔁 Silakan refresh halaman secara manual jika belum berhasil.")

# Layout 2 kolom upload
col1, col2 = st.columns(2)
with col1:
    st.header("📂 Upload Template")
    template_file = st.file_uploader("Upload Template File (.xlsx)", type=["xlsx"], key="template")
with col2:
    st.header("📥 Upload Input Harian")
    input_file = st.file_uploader("Upload Input Data File (.xlsx)", type=["xlsx"], key="input")

# Input tanggal dan nama file hasil
st.divider()
st.header("🗓️ Parameter Proses")
date_input = st.number_input("Pilih Tanggal Input", min_value=1, max_value=31, value=25)
output_filename = st.text_input("Nama File Hasil (tanpa ekstensi)", value="TEMPLATE_FILLED")

# Tampilkan preview otomatis setelah upload
st.divider()
st.header("🔍 Preview Data Upload")
template_df = None
input_data_df = None
if template_file:
    try:
        template_df = pd.read_excel(template_file, sheet_name=0, header=1)
        st.subheader("📑 Template File")
        st.dataframe(template_df.head(), use_container_width=True)
    except Exception as e:
        st.error(f"Gagal membaca file template: {e}")

if input_file:
    try:
        input_data_df = pd.read_excel(input_file, sheet_name=0, header=1)
        st.subheader("📄 Input Harian")
        st.dataframe(input_data_df.head(), use_container_width=True)
    except Exception as e:
        st.error(f"Gagal membaca file input: {e}")

# Tombol proses
st.divider()
st.header("⚙️ Proses Data")
if st.button("🚀 Proses Sekarang") and template_df is not None and input_data_df is not None:
    with st.spinner("Sedang memproses data..."):
        if "Crew Name" not in template_df.columns or "Crew Name" not in input_data_df.columns:
            st.error("❌ Kedua file harus memiliki kolom 'Crew Name'")
        else:
            date_col = None
            for col in template_df.columns:
                if str(date_input) in str(col):
                    date_col = col
                    break

            if date_col is None:
                st.error(f"❌ Tidak ditemukan kolom yang sesuai dengan tanggal {date_input}")
            else:
                template_df[date_col] = template_df[date_col].astype(object)

                for index, row in template_df.iterrows():
                    crew_name = str(row["Crew Name"]).strip()
                    match = input_data_df[input_data_df["Crew Name"].astype(str).str.strip() == crew_name]

                    if not match.empty:
                        val = match[date_col].iloc[0]
                        if pd.notna(val):
                            template_df.at[index, date_col] = val
                        else:
                            template_df.at[index, date_col] = "Na"
                    else:
                        template_df.at[index, date_col] = "Na"

                # Buat workbook
                wb = Workbook()
                ws = wb.active
                ws.title = "Sheet1"
                ws.append(["OPS - REPORT"])

                for r in dataframe_to_rows(template_df, index=False, header=True):
                    ws.append(r)

                output = BytesIO()
                wb.save(output)
                output.seek(0)

                st.success("✅ Data berhasil diproses!")
                st.subheader("📊 Preview Hasil Akhir")
                st.dataframe(template_df.head(), use_container_width=True)

                st.download_button(
                    label="⬇️ Download Hasil sebagai Excel",
                    data=output,
                    file_name=output_filename + ".xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
