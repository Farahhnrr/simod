# SIMOD HPS - Google Apps Script Document Manager

Project ini mengikuti pola `gs_doc_manager`:
- UI HTML Service + `google.script.run` (tanpa endpoint `action` REST)
- Metadata di Google Sheets
- Upload file (base64) ke Google Drive
- Alur: Pendidikan -> HPS -> Upload file -> status `DRAFT/READY`

## Struktur Utama
- `src/WebApp.gs`: entry web app + function yang dipanggil frontend
- `src/SpreadsheetRepo.gs`: operasi Sheet/Drive dan penyimpanan data
- `src/Validation.gs`: helper validasi, sanitasi, mapper row
- `src/Config.gs`: konfigurasi ID, nama sheet, header kolom
- `src/Index.html`: SPA frontend

## Konfigurasi Wajib
Atur salah satu cara berikut:
1. Isi di `src/Config.gs`:
- `STATIC_SHEET_ID`
- `STATIC_DRIVE_FOLDER_ID`

2. Atau isi Script Properties:
- `SHEET_ID`
- `DRIVE_FOLDER_ID`

Keduanya wajib agar web app bisa bootstrap data. Tidak ada pembuatan folder Drive otomatis saat bootstrap, sama seperti pola contoh `gs_doc_manager`.

## Deploy Singkat
```bash
npm install
npm run dev
```

Mode `dev` akan watch perubahan file lokal lalu auto-push ke Apps Script.

Untuk push/deploy manual:
```bash
npm run push
npm run deploy:webapp
```

Web app akan menampilkan 3 langkah:
1. Tambah pendidikan
2. Tambah HPS per pendidikan
3. Upload file HPS dan E-Faktur
