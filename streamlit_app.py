import streamlit as st
import PyPDF2
import camelot
import pandas as pd
import tempfile
import os

# Function to unlock PDF using PdfReader
def unlock_pdf(pdf_file, password):
    try:
        with open(pdf_file, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            
            if reader.is_encrypted:
                try:
                    reader.decrypt(password)
                    st.success("File berhasil didekripsi!")
                except Exception as e:
                    st.error(f"Gagal mendekripsi file. Kesalahan: {e}")
                    return None
            else:
                st.info("File tidak terkunci, tidak perlu dekripsi.")
                return pdf_file  # Kembalikan path file asli jika tidak terkunci

            # Save the unlocked PDF to a temporary file
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_pdf:
                writer = PyPDF2.PdfWriter()
                for page in reader.pages:
                    writer.add_page(page)
                writer.write(temp_pdf)
                return temp_pdf.name
    except Exception as e:
        st.error(f"Error: {e}")
        return None

# Function to extract tables from PDF
def extract_statement(pdf_file):
    try:
        tables = camelot.read_pdf(pdf_file, pages='all', flavor='stream',row_tol=30)
        extracted_data = []

        for table in tables:
            df = table.df
            if df.shape[1] >= 5:  # At least 5 columns
                df.columns = df.iloc[0]  # Use first row as header
                df = df[1:]  # Data starts from the second row
                df.reset_index(drop=True, inplace=True)
                extracted_data.append(df)

        return extracted_data
    except Exception as e:
        st.error(f"Error extracting table: {e}")
        return None

# Function to save data to Excel
def save_to_excel(dataframes):
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_excel:
            with pd.ExcelWriter(temp_excel.name) as writer:
                for i, df in enumerate(dataframes):
                    df.to_excel(writer, sheet_name=f'Page_{i+1}', index=False)
            return temp_excel.name
    except Exception as e:
        st.error(f"Error saving to Excel: {e}")
        return None

# Streamlit UI
st.title("Statement2Sheet")
st.subheader("Aplikasi Untuk Mengubah Tabel Laporan Format PDF ke Format xlsx")
uploaded_file = st.file_uploader("Unggah file PDF", type="pdf")

if uploaded_file is not None:
    try:
        # Simpan file PDF yang di-upload sebagai file sementara
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_file:
            temp_file.write(uploaded_file.getbuffer())
            pdf_path = temp_file.name
        
        # Cek apakah file PDF terkunci
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            if reader.is_encrypted:
                # Jika terkunci, minta password
                password = st.text_input("Masukkan Password", type="password")
                if password:
                    unlocked_pdf = unlock_pdf(pdf_path, password)
                else:
                    st.warning("Masukkan password untuk mendekripsi file.")
                    unlocked_pdf = None
            else:
                unlocked_pdf = pdf_path

        if unlocked_pdf:
            # Tombol untuk konversi ke Excel
            if st.button("Konversi ke Excel"):
                with st.spinner("Mengonversi PDF ke Excel..."):
                    extracted_data = extract_statement(unlocked_pdf)
                    if extracted_data:
                        excel_file = save_to_excel(extracted_data)
                        if excel_file:
                            st.success("Konversi berhasil! File Excel telah dibuat.")

                            # Tombol untuk mengunduh file Excel
                            with open(excel_file, "rb") as f:
                                st.download_button("Unduh File Excel", f, file_name="extracted_statement.xlsx")
    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
    
    finally:
        # Cleanup temporary files
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
        if unlocked_pdf and os.path.exists(unlocked_pdf):
            os.remove(unlocked_pdf)

st.caption('Copyright (c) Muhammad Reza Ubaidillah 2024')
