# Nihaoma Streamlit Dashboard (Google Drive Live Excel)

Dashboard ini membaca file Excel invoice Nihaoma langsung dari Google Drive, lalu menampilkan ringkasan invoice student secara live di web.

## Fitur
- Koneksi ke file Excel `.xlsx` di Google Drive
- Auto refresh berkala supaya perubahan terbaru ikut tampil
- KPI invoice, pembayaran, outstanding, pelunasan, dan pengiriman
- Filter program, status pelunasan, status pengiriman
- Search student / kode invoice
- Detail invoice
- Download CSV hasil filter
- Siap dideploy ke Streamlit Community Cloud

## Struktur workbook yang didukung
Aplikasi ini disesuaikan dengan template Excel Nihaoma yang memiliki sheet:
- `INPUT_DATA`
- `INVOICE_PDF`
- `REKAP_STATUS`
- `SETUP`

Sumber utama yang dibaca dashboard adalah sheet `INPUT_DATA`, sehingga rumus di sheet lain tidak mengganggu.

## Persiapan Google Drive
Supaya dashboard publik bisa membaca file Excel, file di Google Drive harus diatur:
1. Upload workbook Excel ke Google Drive
2. Klik **Share**
3. Ubah akses menjadi **Anyone with the link**
4. Permission minimal **Viewer**
5. Salin link file tersebut

Contoh link:
`https://drive.google.com/file/d/FILE_ID/view?usp=sharing`

Aplikasi bisa memakai:
- full link Google Drive, atau
- langsung `FILE_ID`

## Menjalankan lokal
```bash
python3 -m venv venv
source venv/bin/activate
python3 -m pip install -r requirements.txt
streamlit run app.py
```

## Deploy ke Streamlit Community Cloud
1. Upload semua file project ini ke repository GitHub
2. Buka Streamlit Community Cloud
3. Pilih repo dan deploy `app.py`
4. Setelah itu, isi salah satu secrets / environment variables berikut:

### Opsi A: pakai full link
`GOOGLE_DRIVE_URL = "https://drive.google.com/file/d/FILE_ID/view?usp=sharing"`

### Opsi B: pakai file ID
`GOOGLE_DRIVE_FILE_ID = "FILE_ID"`

### Opsional
`REFRESH_SECONDS = "60"`

## Cara kerja live update
- Setiap interval refresh, aplikasi mengunduh ulang file Excel dari Google Drive
- Jika file Excel di Drive diperbarui, dashboard akan membaca versi terbaru
- Jadi update bersifat **near real-time**, tergantung interval refresh

## Catatan penting
- Jika file **tidak public/shared**, dashboard publik tidak akan bisa membaca file
- Excel cocok bila admin yang mengedit hanya sedikit
- Jika nanti user makin banyak dan butuh edit bersamaan, sebaiknya migrasi ke Google Sheets atau database

## File yang ada
- `app.py` → aplikasi dashboard Streamlit
- `requirements.txt` → dependency Python
- `README.md` → panduan setup dan deploy

## Tips
Kalau Anda ingin dashboard benar-benar publik dan stabil:
- simpan workbook final di satu file Google Drive pusat
- jangan terlalu sering ganti nama / pindahkan file
- gunakan satu admin flow untuk update data
