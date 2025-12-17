# Karyamas Plantation – Aplikasi Transkrip Nilai (Offline-first)

## 1) Frontend (HTML)
Buka `index.html` (bisa via hosting lokal / GitHub Pages).

**Library:** Bootstrap 5, idb (IndexedDB), SheetJS (XLSX), jsPDF + AutoTable.

### Konfigurasi
1. Deploy GAS WebApp (lihat bagian GAS di bawah) dan copy URL WebApp.
2. Login sebagai `admin` / `123456`.
3. Buka menu **Setting (Admin)** → paste URL pada kolom `Google Apps Script WebApp URL`.

## 2) Google Apps Script (Backend)
1. Buat Google Spreadsheet baru untuk database.
2. Buat Apps Script baru, paste isi `Code.gs`.
3. Isi `SPREADSHEET_ID` dengan ID spreadsheet.
4. Deploy: **Deploy → New deployment → Web app**
   - Execute as: Me
   - Who has access: Anyone with the link
5. Copy WebApp URL dan tempelkan ke frontend.

### Sheet yang digunakan
- users
- master_peserta
- master_materi
- master_pelatihan
- master_bobot
- master_predikat
- nilai

Script akan otomatis membuat sheet dan header jika belum ada.

## 3) Pola Kerja Offline
- Input nilai selalu tersimpan ke IndexedDB dan masuk **queue**.
- Saat online, klik **Sync** (atau otomatis ketika koneksi kembali) untuk kirim ke Google Sheet.

## 4) Export
- **Export XLSX** dari filter yang aktif.
- **Export Transkrip PDF**: berdasarkan data `test_type=Final` dari filter.

Catatan:
Template PDF masih versi 1 (layout mirip contoh, bisa disempurnakan agar benar-benar identik).
