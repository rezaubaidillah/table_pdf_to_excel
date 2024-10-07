# Statement2Sheet

Statement2Sheet adalah aplikasi web yang dibangun dengan Streamlit untuk mengubah tabel laporan dalam format PDF berpassword menjadi format Excel (.xlsx). Aplikasi ini menggunakan library Python seperti PyPDF2 untuk membuka PDF yang terpassword dan Camelot untuk mengekstrak tabel dari file PDF.

## Fitur

- **Mendukung file PDF berpassword**: Aplikasi ini dapat membuka dan mendekripsi file PDF yang dilindungi dengan password.
- **Ekstraksi tabel**: Menyediakan fitur untuk mengekstrak tabel dari file PDF dan menyimpannya ke dalam format Excel.
- **Antarmuka pengguna yang sederhana**: Antarmuka intuitif yang memudahkan pengguna untuk mengunggah file dan melakukan konversi.

## Prasyarat

Sebelum menjalankan aplikasi ini, pastikan Anda telah menginstal paket-paket berikut:

- Python 3.6+
- Streamlit
- PyPDF2
- Camelot
- Pandas
- OpenPyXL

Anda dapat menginstal semua dependensi dengan menjalankan perintah berikut:

```bash
pip install streamlit PyPDF2 camelot-py pandas openpyxl
