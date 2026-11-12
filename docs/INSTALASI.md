# Instalasi: membuat Google Sheets ROTI RONA

> Tidak ada password, API key, atau token yang diminta. Script hanya meminta izin Google standar (membuat & mengedit spreadsheet **milik akun Anda sendiri**) saat pertama kali dijalankan.

## Cara A — Google Apps Script (disarankan, menghasilkan Google Sheets asli + grafik native)

1. Buka <https://script.google.com> → **Proyek baru**.
2. Hapus isi `Code.gs` bawaan, lalu salin-tempel seluruh isi file **`Code.gs`** dari repositori ini.
3. (Opsional) **Project Settings → Show "appsscript.json"** lalu samakan dengan `appsscript.json` di repositori (zona waktu `Asia/Jakarta`, runtime V8).
4. Pilih fungsi **`setup`** pada dropdown di toolbar → klik **Run**.
5. Saat diminta, **Review permissions** → pilih akun Google Anda → **Advanced → Go to … (unsafe)** → **Allow**.
   (Peringatan muncul karena script buatan sendiri belum diverifikasi Google — normal.)
6. Tunggu ±20–40 detik. Buka **View → Logs** (atau *Execution log*) untuk melihat URL spreadsheet baru bernama **“ROTI RONA - Pembukuan & Dashboard Keuangan”**; file juga muncul di Google Drive Anda.

Variasi: jika script dibuat dari **Extensions → Apps Script** pada Google Sheets kosong, jalankan **`setupDiSpreadsheetIni`**. Fungsi ini menolak berjalan bila sheet DASHBOARD/TRANSAKSI/KATEGORI/PETUNJUK sudah ada (agar data tidak tertimpa).

## Cara B — Impor file Excel (alternatif, tanpa script)

1. Buka <https://drive.google.com> → **New → File upload** → pilih `ROTI_RONA_Pembukuan.xlsx`.
2. Klik kanan file → **Open with → Google Sheets**, lalu **File → Save as Google Sheets**.
3. **File → Settings** → set *Locale* ke **Indonesia** dan *Time zone* ke **(GMT+07:00) Jakarta** agar format tanggal/Rupiah mengikuti Indonesia.
4. Periksa grafik di DASHBOARD. Pada beberapa kasus grafik hasil impor perlu disesuaikan ulang (Cara A lebih rapi).

## Bagian yang perlu Anda konfigurasi

| Item | Lokasi | Keterangan |
|---|---|---|
| **Saldo Awal** | DASHBOARD sel `I5` | Kas sebelum transaksi pertama dicatat. Default Rp 15.000.000 (simulasi). |
| **Tahun & Bulan laporan** | DASHBOARD `C5` dan `E5` | Dropdown. |
| **Kategori baru** | KATEGORI kolom B–C | Isi baris kosong; langsung muncul di dropdown. |
| **Data demo → data asli** | Menu **ROTI RONA → Hapus data simulasi** | Muncul setelah file dibuka ulang (fungsi `onOpen`). Atau kosongkan manual. |
| **Daftar tahun filter** | KATEGORI kolom I | Default 2024–2035; ubah bila perlu. |
| **Nama usaha** | Teks judul di DASHBOARD/TRANSAKSI | Ganti “ROTI RONA” dengan nama usaha Anda. |

## Mengimpor `sample_transactions.csv` (opsional)

Kolom CSV: `Tanggal, Jenis Transaksi, Kategori, Keterangan, Metode Pembayaran, Nominal`. Tempel ke sheet TRANSAKSI mulai sel **A5** dengan memperhatikan bahwa kolom **B, H, I, J** berisi formula (jangan ditimpa): tempel kolom CSV 1 → A, kolom 2–5 → C–F, kolom 6 → G (*Edit → Paste special → Values only*).

## Pemecahan masalah

- **“Authorization required” berulang** — pastikan Anda login dengan satu akun Google saja.
- **Grafik kosong** — pastikan Tahun di DASHBOARD sesuai tahun data; grafik bulanan menampilkan 12 bulan pada tahun terpilih.
- **Kolom “Transaksi perlu dicek” > 0** — filter kolom *Cek Data* di TRANSAKSI untuk nilai selain “OK”.
