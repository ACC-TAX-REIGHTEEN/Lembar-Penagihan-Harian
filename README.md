# 📋 Ambil AR — Lembar Tagihan & Sinkronisasi Google Sheets

> **Satu klik: ekspor AR Accurate → pembersihan multi-tahap → lembar tagihan per penagih siap cetak + inject ke Google Sheets**

Pipeline Python empat belas langkah yang membaca ekspor daftar piutang dari Accurate pilihan: Daftar Piutang Penjualan atau Piutang Persales (`Piutang.xls`) dan data Giro wajib jika menggunakan Piutang Persales (`Giro.xls`), menjalankan serangkaian validasi dan pembersihan data (hapus piutang lunas via Giro, hapus saldo nol, sinkronisasi nama dari master data, filter faktur pending), mengelompokkan tagihan per **penagih/sales**, menghasilkan **`Print_AR.xlsm`** — lembar tagihan per penagih berformat template resmi siap cetak — sekaligus menyuntikkan seluruh data ke **Google Sheets** dan mencetak otomatis ke printer yang dipilih.

---

## 📋 Daftar Isi

- [Gambaran Umum](#-gambaran-umum)
- [Fitur Utama](#-fitur-utama)
- [Prasyarat](#-prasyarat)
- [Struktur Folder & File](#-struktur-folder--file)
- [Cara Penggunaan](#-cara-penggunaan)
- [Alur Kerja Pipeline](#-alur-kerja-pipeline)
- [Detail Tiap Skrip](#-detail-tiap-skrip)
- [Konfigurasi `piutang.conf`](#-konfigurasi-piutangconf)
- [Setup Google Sheets API](#-setup-google-sheets-api)
- [Output](#-output)
- [Folder VBA — Alternatif Cetak via Macro](#-folder-vba--alternatif-cetak-via-macro)
- [Troubleshooting](#-troubleshooting)
- [Catatan Penting](#-catatan-penting)

---

## 🗂️ Gambaran Umum

`Ambil AR` mengotomasi seluruh alur kerja harian admin piutang: dari ekspor Accurate hingga lembar tagihan siap dibagikan ke penagih. Versi terbaru menambahkan beberapa lapisan validasi data sebelum data difilter per penagih, sehingga lembar tagihan yang dihasilkan sudah bersih dari:

- Faktur yang telah terbayar lunas via Giro
- Faktur dengan saldo nol
- Faktur yang sedang dalam status "pending" (belum diselesaikan)

Pipeline juga dapat menyeragamkan nama pelanggan secara otomatis dari spreadsheet master, serta mencetak langsung ke printer tanpa membuka Excel secara manual.

---

## ✨ Fitur Utama

- **Auto-pembersihan piutang lunas via Giro** — Faktur yang nilai pembayaran gironya cocok dengan nilai faktur (dalam batas toleransi `giro_cut`) otomatis dihapus dari daftar tagihan.
- **Filter faktur pending** — Faktur yang terdaftar di sheet Pending tapi belum memiliki tanggal penyelesaian otomatis dikeluarkan dari tagihan.
- **Sinkronisasi nama pelanggan dari master data** — Nama pelanggan di data AR dapat diperbarui secara otomatis dari spreadsheet master menggunakan kode pelanggan sebagai kunci.
- **Hapus baris saldo nol** — Faktur dengan `Sisa Piutang = 0` dihapus otomatis setelah sinkronisasi master.
- **Dua mode generate template** — Gunakan `xlwings` (butuh Excel terinstall) atau **Pure Python** via `openpyxl` (tidak butuh Excel) sesuai kondisi sistem.
- **Cetak otomatis via Python** — Dialog pemilihan printer muncul otomatis dan mencetak setiap blok laporan penagih secara terpisah (Landscape, fit to 1 page, margin 0.25 inch).
- **Mapping penagih dari konfigurasi** — Setiap kode pelanggan dipetakan ke nama penagih melalui `piutang.conf`. Tidak perlu mengubah kode; cukup edit konfigurasi.
- **Rekalkulasi umur piutang** — Kolom `Umur JT` dihitung ulang berdasarkan selisih `Tgl Faktur` dan hari ini, sehingga nilai selalu akurat.
- **Kolom Terbayar otomatis** — Menghitung `Nilai Faktur − Sisa Piutang` per baris; dikosongkan jika nol.
- **Template `.xlsm` dengan style & shape tersalin** — Header, baris data, total `=SUM()`, footer TTD, dan gambar/logo tersalin dari template.
- **Inject ke Google Sheets** — 14 kolom per faktur disisipkan sebelum baris terakhir sheet target dengan mewarisi format bawaan.
- **Semua fitur kondisional** — Giro, Master, Pending, Pure Python, dan Print masing-masing punya flag aktif/nonaktif di `piutang.conf`; skip otomatis jika tidak diaktifkan.

---

## 🔧 Prasyarat

### Python
Python **3.8+** disarankan.

### Library yang dibutuhkan

```bash
pip install pandas openpyxl xlrd xlsxwriter xlwings gspread google-auth requests pywin32
```

| Library | Digunakan di | Kegunaan |
|---|---|---|
| `pandas` | Skrip 1–2P | Baca `.xls`, bersihkan, filter, merge, groupby |
| `xlsxwriter` | Skrip 1, 2 | Buat `.xlsx` sementara dengan format angka |
| `xlrd` | Skrip 1 | Baca file legacy `.xls` dari Accurate dan Giro |
| `openpyxl` | Skrip 1D, 3P, 4 | Hapus baris, generate template Pure Python, unmerge |
| `xlwings` | Skrip 3, 6 | Tulis ke `.xlsm` via Excel COM; cetak via Excel API |
| `requests` | Skrip 1B, 1E | Unduh file dari Google Sheets |
| `gspread` | Skrip 5 | Klien Google Sheets API |
| `google-auth` | Skrip 5 | Autentikasi via Service Account |
| `pywin32` | Skrip 6 | Enumerasi dan pemilihan printer (Windows only) |
| `tkinter` | Skrip 6 | Dialog pemilihan printer (termasuk di Python stdlib) |
| `configparser`, `re`, `os`, `glob`, `shutil`, `subprocess`, `sys`, `datetime`, `pathlib`, `ctypes`, `copy` | Semua | Standard library |

### Aplikasi wajib / kondisional

| Aplikasi | Diperlukan untuk | Kondisi |
|---|---|---|
| **Microsoft Excel** | Skrip 3 (xlwings), Skrip 6 (cetak) | Wajib jika `pr_process = No`; Skrip 6 selalu butuh Excel |
| **Windows OS** | Skrip 6 (pywin32, ctypes) | Skrip 6 hanya berjalan di Windows |

> **Tanpa Excel:** Aktifkan `[PURE] pr_process = Ya` di `piutang.conf`. Skrip `3_CalculateARPurePython.py` menggunakan `openpyxl` murni. Namun `3_CalculateAR.py` (xlwings) tetap dijalankan lebih dulu oleh orkestrator — jika gagal karena tidak ada Excel, pipeline akan berhenti. Untuk sistem tanpa Excel, jalankan `3_CalculateARPurePython.py` secara mandiri setelah step sebelumnya selesai.
>
> **Catatan `xlrd`:** Gunakan versi yang kompatibel dengan `.xls`:
> ```bash
> pip install "xlrd>=1.0.0,<2.0.0"
> ```

---

## 📁 Struktur Folder & File

```
📦 Lembar-Penagihan-Harian/
│
├── 📄 Ambil AR.py                       ← Orkestrator utama. Jalankan ini
│
├── 📄 Piutang.xls                       ← [INPUT] Ekspor piutang dari Accurate (wajib)
├── 📄 Giro.xls                          ← [INPUT] Rekap Giro/cek (opsional)
├── 📄 Ekspor Data.png                   ← Panduan visual cara ekspor dari Accurate
│
├── 📁 Dapur/                            ← Folder pipeline (jangan diubah)
│   ├── 📄 __init__.py
│   ├── 📄 1_CleanerAcc.py              ← Bersihkan Piutang.xls → Piutang_clean_temp.xlsx
│   ├── 📄 1_CleanerAccGiro.py          ← ✨ Bersihkan Giro.xls → Giro_temp.xlsx
│   ├── 📄 1B_DownloaderMasterData.py   ← ✨ Unduh Master_temp.xlsx dari Google Sheets
│   ├── 📄 1C_MergedMaster2Main.py      ← ✨ Sinkronkan nama pelanggan dari master
│   ├── 📄 1D_CleanZeroAR.py            ← ✨ Hapus baris dengan Sisa Piutang = 0
│   ├── 📄 1E_DownloaderPendingData.py  ← ✨ Unduh Pending_temp.xlsx dari Google Sheets
│   ├── 📄 2_CompareGiro.py             ← ✨ Hapus faktur yang sudah lunas via Giro
│   ├── 📄 2_ComparePending.py          ← ✨ Hapus faktur yang sedang pending
│   ├── 📄 2_FilterAR.py               ← Filter per penagih + hitung Terbayar & total
│   ├── 📄 3_CalculateAR.py            ← Susun ke TEMPLATE.xlsm → Print_AR.xlsm (xlwings)
│   ├── 📄 3_CalculateARPurePython.py  ← ✨ Alternatif Pure Python (openpyxl, tanpa Excel)
│   ├── 📄 4_HelperCleaningData.py     ← Ratakan merge, isi nama penagih, hapus non-data
│   ├── 📄 5_InjectDataToSS.py         ← Sisipkan 14 kolom ke Google Sheets
│   ├── 📄 6_PrintByPython.py          ← ✨ Cetak otomatis ke printer via dialog (Windows)
│   ├── 📄 TEMPLATE.xlsm               ← Template lembar tagihan (jangan dihapus)
│   ├── 📄 credentials.json            ← Kredensial Google Service Account (rahasia!)
│   └── 📄 piutang.conf                ← Semua konfigurasi pipeline
│
└── 📁 VBA/                             ← ✨ Alternatif macro Excel untuk cetak
    └── 📄 Print.bas                    ← VBA macro: cetak per blok penagih
```

> ✨ = File baru yang ditambahkan pada versi ini.

---

## 🚀 Cara Penggunaan

### Langkah 1 — Siapkan file input

Letakkan di folder utama:
- **`Piutang.xls`** (wajib) — ekspor laporan piutang dari Accurate
- **`Giro.xls`** (opsional) — rekap penerimaan Giro/cek dari tim keuangan

> Lihat `Ekspor Data.png` untuk panduan visual cara mengekspor dari Accurate.

### Langkah 2 — Sesuaikan `piutang.conf`

Buka `Dapur/piutang.conf` dan perbarui:
- Mapping `[NAMA SALES]` + `[KODE PELANGGAN]`
- Metadata: `[PERUSAHAAN]`, `[DIVISI]`, `[TANGGAL]`, `[INPUT]`
- Aktifkan/nonaktifkan fitur: `[GIRO]`, `[MASTER]`, `[PENDING]`, `[PURE]`, `[PPRINT]`
- URL Google Sheets di seksi `[SS]`, `[MASTER]`, `[PENDING]`

Lihat panduan lengkap di [Konfigurasi `piutang.conf`](#-konfigurasi-piutangconf).

### Langkah 3 — Pasang kredensial

Ganti isi `Dapur/credentials.json` dengan file JSON Google Service Account Anda. Lihat [Setup Google Sheets API](#-setup-google-sheets-api).

### Langkah 4 — Jalankan

```bash
python "Ambil AR.py"
```

### Langkah 5 — Pantau progress

```
--> Memulai eksekusi pembersihan data utama
--> SUKSES! File tersimpan rapi di: Piutang_clean_temp.xlsx
--> Memulai eksekusi pembersihan data utama (Giro)
--> File Giro_temp.xlsx telah berhasil dibuat
--> Memulai pengunduhan data master
--> Berhasil! File disimpan sebagai Master_temp.xlsx
--> Memulai penggabungan data master ke data utama
--> Selesai! Berhasil memperbarui 24 baris data 'Nama Pelanggan'
--> Memulai pembersihan saldo piutang nol
--> Berhasil menghapus 3 baris dengan Sisa Piutang 0.
--> Memulai pengunduhan data pendingan
--> Memulai komparasi data giro
--> Berhasil menghapus 5 data yang klop dari 'Piutang_clean_temp.xlsx'
--> Memulai komparasi data pendingan
--> Memulai eksekusi filter data sementara
--> Memulai eksekusi menyalin dan menyusun data pada template
--> Menggunakan metode pure code yang ringan
--> Memulai persiapan data untuk disusun ke Spreadsheets
--> Memulai unggah data ke Spreadsheets
--> Memulai menjalankan print otomatis dengan konfigurasi
--> [Dialog pemilihan printer muncul]
--> Selesai! Total ada 3 kelompok laporan yang dicetak.
--> Semua proses telah selesai dijalankan.
```

---

## 🔄 Alur Kerja Pipeline

```
[Mulai: Ambil AR.py]
   │
   ├─── Validasi: folder Dapur/ + 17 file syarat ada
   ├─── Validasi: Piutang.xls ada di folder utama
   ├─── Bersihkan Dapur/: *temp.xlsx, Giro.xls, Piutang.xls lama
   ├─── Salin Piutang.xls → Dapur/ (wajib)
   ├─── Salin Giro.xls → Dapur/ (jika ada)
   │
   ├─── [1] 1_CleanerAcc.py
   │       Scan 150 baris → deteksi header → pilih 7 kolom
   │       Bersihkan angka, buang baris NaN/header
   │       → Piutang_clean_temp.xlsx
   │
   ├─── [1G] 1_CleanerAccGiro.py  (skip jika giro_stats ≠ Ya)
   │       Baca Giro.xls → filter baris ≥ 9 kolom tidak kosong
   │       → Giro_temp.xlsx
   │
   ├─── [1B] 1B_DownloaderMasterData.py  (skip jika masdatus ≠ Ya)
   │       Unduh spreadsheet [MASTER] url → Master_temp.xlsx
   │
   ├─── [1C] 1C_MergedMaster2Main.py  (skip jika masdatus ≠ Ya)
   │       Baca Master_temp.xlsx → bangun lookup {mas_col_key: mas_col_ret}
   │       Update kolom Nama Pelanggan di Piutang_clean_temp.xlsx
   │       → Piutang_clean_temp.xlsx (diperbarui)
   │
   ├─── [1D] 1D_CleanZeroAR.py
   │       Hapus baris di Piutang_clean_temp.xlsx mana Sisa Piutang = 0
   │       → Piutang_clean_temp.xlsx (diperbarui)
   │
   ├─── [1E] 1E_DownloaderPendingData.py  (skip jika pend_stats ≠ Ya)
   │       Unduh spreadsheet [PENDING] pend_url → Pending_temp.xlsx
   │
   ├─── [2G] 2_CompareGiro.py  (skip jika giro_stats ≠ Ya)
   │       Gabungkan Piutang vs Giro per (Kode Pelanggan, No. Faktur)
   │       Hapus baris: |Nilai Faktur - Total Giro| ≤ giro_cut
   │       → Piutang_clean_temp.xlsx (diperbarui)
   │
   ├─── [2P] 2_ComparePending.py  (skip jika pend_stats ≠ Ya)
   │       Cari No. Faktur di Pending yang pend_col_ret KOSONG
   │       Hapus baris tersebut dari Piutang_clean_temp.xlsx
   │       → Piutang_clean_temp.xlsx (diperbarui)
   │
   ├─── [2F] 2_FilterAR.py
   │       Bangun map {Kode: Penagih} dari piutang.conf
   │       Filter → recalkulasi Umur JT → hitung Terbayar
   │       Urutkan + sisipkan baris TOTAL per penagih
   │       → Laporan_Piutang_Penagih_temp.xlsx
   │
   ├─── [3] 3_CalculateAR.py  (via xlwings, butuh Excel)
   │       Buka TEMPLATE.xlsm via Excel COM
   │       Loop per penagih: salin header+data+total+footer+shapes
   │       → Print_AR.xlsm
   │       → DISALIN ke folder utama segera setelah step ini
   │
   ├─── [3P] 3_CalculateARPurePython.py  (skip jika pr_process ≠ Ya)
   │       Buka TEMPLATE.xlsm via openpyxl (tanpa Excel)
   │       Salin cell style, border, fill, height, formula
   │       → Print_AR.xlsm (menimpa hasil xlwings jika pr_process = Ya)
   │
   ├─── [4] 4_HelperCleaningData.py
   │       Unmerge sel, isi nama penagih di kolom A
   │       Hapus baris non-data → ratakan tinggi baris
   │       → Print_AR_temp.xlsx
   │
   ├─── [5] 5_InjectDataToSS.py
   │       Susun 14 kolom → autentikasi → sisipkan sebelum baris terakhir
   │       → Google Sheets diperbarui
   │
   ├─── [6] 6_PrintByPython.py  (skip jika [PPRINT] status ≠ Ya, Windows only)
   │       Tampilkan dialog pemilihan printer (tkinter)
   │       Buka Print_AR.xlsm via xlwings
   │       Deteksi blok "LAPORAN HASIL TAGIHAN" → "TTD SALES & COLLECTOR"
   │       Cetak tiap blok: Landscape, FitToPage, margin 0.25", Paper Letter
   │
   └─── Cleanup: hapus *temp.xlsx, Giro.xls, Piutang.xls, Print_AR.xlsm dari Dapur/
```

---

## 🔍 Detail Tiap Skrip

### Skrip 1 — `1_CleanerAcc.py`
Membaca `Piutang.xls` tanpa asumsi posisi header, scan 150 baris pertama, pilih 7 kolom berdasarkan nama, bersihkan angka format lokal. Output: `Piutang_clean_temp.xlsx`.

---

### Skrip 1G — `1_CleanerAccGiro.py` ✨
Membersihkan `Giro.xls` menjadi `Giro_temp.xlsx`. Baris dianggap valid jika memiliki ≥ 9 kolom tidak kosong (mengabaikan baris total/header yang sparse). Kolom `Total Diterima` dan `Nilai terima` dikonversi ke numerik. **Skip otomatis** jika `[GIRO] giro_stats ≠ Ya`.

---

### Skrip 1B — `1B_DownloaderMasterData.py` ✨
Mengunduh file master pelanggan dari Google Sheets ke `Master_temp.xlsx`. URL diambil dari `[MASTER] url` di `piutang.conf`. **Skip otomatis** jika `masdatus ≠ Ya` atau URL kosong.

---

### Skrip 1C — `1C_MergedMaster2Main.py` ✨
Membaca `Master_temp.xlsx` dan membangun lookup `{mas_col_key → mas_col_ret}`. Setiap baris di `Piutang_clean_temp.xlsx` yang kode pelanggannya ditemukan di master akan diperbarui kolom `Nama Pelanggan`-nya. Berguna jika nama di Accurate tidak konsisten dengan nama resmi di sistem master. **Skip otomatis** jika `masdatus ≠ Ya`.

---

### Skrip 1D — `1D_CleanZeroAR.py` ✨
Menghapus semua baris di `Piutang_clean_temp.xlsx` di mana nilai kolom `Sisa Piutang` adalah `0`. Berjalan setiap kali pipeline dieksekusi tanpa kondisi apapun.

---

### Skrip 1E — `1E_DownloaderPendingData.py` ✨
Mengunduh data faktur pending dari Google Sheets ke `Pending_temp.xlsx`. URL dari `[PENDING] pend_url`. **Skip otomatis** jika `pend_stats ≠ Ya` atau URL kosong.

---

### Skrip 2G — `2_CompareGiro.py` ✨
Mencocokkan data AR dengan data Giro. Pencocokan dilakukan per kombinasi `(Kode Pelanggan ↔ No. Pelanggan Giro, No. Faktur ↔ No. Faktur SO)`. Jumlah Giro diagregasi per faktur, lalu dibandingkan dengan `Nilai Faktur`. Jika `|Nilai Faktur - Total Giro| ≤ giro_cut` → baris dianggap **lunas** dan dihapus dari `Piutang_clean_temp.xlsx`. **Skip otomatis** jika `giro_stats ≠ Ya`.

---

### Skrip 2P — `2_ComparePending.py` ✨
Membaca sheet Pending dari `Pending_temp.xlsx`. Faktur yang ada di kolom `pend_col_key` namun kolom `pend_col_ret` (biasanya tanggal penyelesaian)-nya **kosong/NaT** → berarti masih pending → dihapus dari `Piutang_clean_temp.xlsx`. **Skip otomatis** jika `pend_stats ≠ Ya`.

---

### Skrip 2F — `2_FilterAR.py`
Membangun mapping `{kode: penagih}` dari `piutang.conf`, memfilter data, merecalkulasi Umur JT, menghitung Terbayar, mengurutkan, dan menyisipkan baris `TOTAL [Penagih]`.

---

### Skrip 3 — `3_CalculateAR.py`
Menyusun lembar cetak via `xlwings` (Excel COM). Menyalin header, baris data, baris total `=SUM()`, footer TTD, dan shapes/logo dari `TEMPLATE.xlsm`. **Butuh Microsoft Excel terinstall.**

---

### Skrip 3P — `3_CalculateARPurePython.py` ✨
Alternatif Skrip 3 menggunakan `openpyxl` murni tanpa Excel. Menyalin font, border, fill, alignment, tinggi baris, dan formula dari template. Jika aktif (`pr_process = Ya`), hasilnya **menimpa** output Skrip 3. Cocok untuk lingkungan server/Linux tanpa Microsoft Excel.

---

### Skrip 4 — `4_HelperCleaningData.py`
Membuka `Print_AR.xlsm`, unmerge semua sel, isi kolom A dengan nama penagih per baris data, hapus baris non-data, ratakan tinggi baris. Output: `Print_AR_temp.xlsx` siap inject ke Sheets.

---

### Skrip 5 — `5_InjectDataToSS.py`
Menyuntikkan 14 kolom per faktur ke Google Sheets. Pencarian sheet case-insensitive. Disisipkan sebelum baris terakhir dengan `inherit_from_before=True`.

---

### Skrip 6 — `6_PrintByPython.py` ✨ (Windows only)
Menampilkan dialog pemilihan printer berbasis `tkinter`, membuka `Print_AR.xlsm` via `xlwings`, lalu memindai blok laporan dari penanda `"LAPORAN HASIL TAGIHAN"` hingga `"TTD SALES & COLLECTOR"`. Setiap blok dicetak secara terpisah dengan pengaturan halaman:
- **Orientasi:** Landscape
- **Kertas:** Letter
- **Zoom:** Fit to 1 page wide × 1 page tall
- **Margin:** 0.25 inch di semua sisi

**Skip otomatis** jika `[PPRINT] status ≠ Ya`. **Hanya berjalan di Windows** karena menggunakan `pywin32` dan `ctypes`.

---

## ⚙️ Konfigurasi `piutang.conf`

### Mapping penagih & pelanggan

```ini
[NAMA SALES]
DSR - Kristia Devi

[KODE PELANGGAN]
PW-2063 (2)
PW-2048
KOSONG

[NAMA SALES]
C - Samsul Aziz

[KODE PELANGGAN]
PWT-3158
PWT-2068
```

Setiap `[NAMA SALES]` diikuti langsung oleh `[KODE PELANGGAN]`. Kode yang tidak terdaftar diabaikan oleh Skrip 2F.

---

### Metadata laporan

```ini
[PERUSAHAAN]
PTM

[DIVISI]
PCMO

[TANGGAL]
03/08/2026

[INPUT]
FEBIKA
```

---

### `[GIRO]` — Pembersihan faktur lunas via Giro ✨

```ini
[GIRO]
giro_stats = Ya     ; Aktifkan: Ya | Nonaktifkan: No
giro_cut = 1000     ; Toleransi selisih (Rupiah). Faktur dianggap lunas jika
                    ; |Nilai Faktur - Total Giro| ≤ nilai ini
```

| Key | Keterangan |
|---|---|
| `giro_stats` | `Ya` → aktifkan pembersihan via Giro (butuh `Giro.xls` di folder utama) |
| `giro_cut` | Batas toleransi selisih dalam Rupiah. `0` = harus sama persis |

---

### `[MASTER]` — Sinkronisasi nama pelanggan ✨

```ini
[MASTER]
masdatus = Ya               ; Aktifkan: Ya | Nonaktifkan: No
url = https://docs.google.com/spreadsheets/d/ID/edit
mas_sheet = Customer PCMO   ; Nama sheet di Master spreadsheet
mas_col_key = NOPEL         ; Kolom kunci (kode pelanggan) di master
mas_col_ret = NAMA PELANGGAN ; Kolom nilai (nama resmi) yang akan disalin
```

---

### `[PENDING]` — Filter faktur pending ✨

```ini
[PENDING]
pend_stats = Ya             ; Aktifkan: Ya | Nonaktifkan: No
pend_url = https://docs.google.com/spreadsheets/d/ID/edit
pend_sheet = NamaSheet      ; Sheet yang berisi data pending
pend_col_key = No. Faktur   ; Kolom nomor faktur di sheet pending
pend_col_ret = Tgl Selesai  ; Kolom tanggal penyelesaian (kosong = masih pending)
```

Faktur yang ada di `pend_col_key` namun `pend_col_ret`-nya kosong/NaT akan dihapus dari daftar tagihan.

---

### `[PURE]` — Mode generate template tanpa Excel ✨

```ini
[PURE]
pr_process = Ya    ; Ya → gunakan 3_CalculateARPurePython.py (openpyxl, tanpa Excel)
                   ; No → gunakan 3_CalculateAR.py saja (xlwings, butuh Excel)
```

Jika `Ya`, output Skrip 3P menimpa output Skrip 3.

---

### `[PPRINT]` — Cetak otomatis via Python ✨

```ini
[PPRINT]
status = No    ; Ya → tampilkan dialog printer + cetak Print_AR.xlsm (Windows only)
               ; No → skip, tidak ada pencetakan otomatis
```

---

### `[SS]` — Google Sheets target (inject data)

```ini
[SS]
url = https://docs.google.com/spreadsheets/d/ID/edit
sheet_name = LPH PCMO 2026
```

---

## 🔑 Setup Google Sheets API

### 1. Buat Service Account

1. Buka [Google Cloud Console](https://console.cloud.google.com/) → buat/pilih project.
2. Aktifkan **Google Sheets API** dan **Google Drive API**.
3. Buka **IAM & Admin → Service Accounts** → buat Service Account baru.
4. Di tab **Keys** → buat key baru tipe **JSON** → file terunduh otomatis.

### 2. Pasang kredensial

Ganti isi `Dapur/credentials.json` dengan file JSON yang diunduh.

### 3. Berikan akses ke semua spreadsheet

Tambahkan `client_email` dari `credentials.json` sebagai **Editor** di:
- Google Sheets target inject data (`[SS] url`)
- Google Sheets master pelanggan (`[MASTER] url`) — jika `masdatus = Ya`
- Google Sheets data pending (`[PENDING] pend_url`) — jika `pend_stats = Ya`

---

## 📤 Output

### 1. `Print_AR.xlsm` — Lembar tagihan siap cetak

Satu blok per penagih secara vertikal, masing-masing berisi:
- **Header** — nama perusahaan, divisi, tanggal, nama penagih, penginput
- **Tabel faktur** — No., Kode, Nama Pelanggan, Umur JT, No. Faktur, Tgl Faktur, Nilai Faktur, Terbayar, Sisa Piutang
- **Baris TOTAL TAGIHAN** — formula `=SUM()` otomatis
- **Footer TTD** — area tanda tangan Sales & Collector

### 2. Baris baru di Google Sheets (14 kolom)

| # | Kolom | Sumber |
|---|---|---|
| 1 | Perusahaan | `[PERUSAHAAN]` |
| 2 | Nama Penagih | Dari mapping |
| 3 | Divisi | `[DIVISI]` |
| 4 | Tanggal | `[TANGGAL]` |
| 5 | Input | `[INPUT]` |
| 6 | No. | Nomor urut per penagih |
| 7 | Kode | Kode pelanggan |
| 8 | Nama Pelanggan | Nama (setelah sinkronisasi master jika aktif) |
| 9 | Umur JT | Dihitung ulang dari hari ini |
| 10 | No. Faktur | Nomor faktur |
| 11 | Tgl Faktur | DD/MM/YYYY |
| 12 | Nilai Faktur | Nilai faktur asli |
| 13 | Terbayar | Nilai Faktur − Sisa Piutang |
| 14 | Sisa Piutang | Sisa piutang saat ini |

---

## 📁 Folder VBA — Alternatif Cetak via Macro

Folder `VBA/` berisi `Print.bas` — macro VBA yang dapat diimpor ke `TEMPLATE.xlsm` sebagai alternatif dari `6_PrintByPython.py`. Keduanya melakukan hal yang sama: mencetak setiap blok laporan (dari `LAPORAN HASIL TAGIHAN` hingga `TTD SALES & COLLECTOR`) secara terpisah ke printer yang dipilih.

**Cara import ke Excel:**
1. Buka `Print_AR.xlsm` di Excel
2. Tekan `Alt + F11` → Visual Basic Editor
3. File → Import File → pilih `Print.bas`
4. Jalankan `CetakLaporanARPerBlok` dari menu Macro (Alt + F8)

**Kapan gunakan VBA daripada Python:**
- Jika ingin menjalankan cetak manual dari dalam Excel tanpa perlu pipeline Python
- Jika `[PPRINT] status = No` (Python skip) tapi tetap butuh cetak terstruktur per blok

---

## 🛠️ Troubleshooting

### ❌ `File Piutang.xls tidak ditemukan untuk diproses`
Pastikan file ada di folder utama dengan nama **persis** `Piutang.xls`.

### ❌ `File 1_CleanerAccGiro.py tidak ditemukan di dalam folder Dapur`
File ini wajib ada meski Giro tidak digunakan — orkestrator memvalidasinya sebelum memulai. Pastikan semua 17 file syarat ada di `Dapur/`.

### ❌ Skrip 1_CleanerAccGiro gagal tapi saya tidak punya Giro.xls
Jika `giro_stats = Ya` di `piutang.conf` tapi `Giro.xls` tidak ada di folder utama, skrip akan error. Solusi: ubah ke `giro_stats = No`, atau sediakan `Giro.xls`.

### ❌ `1C_MergedMaster2Main.py: Kolom tidak ditemukan`
Nilai `mas_col_key` atau `mas_col_ret` di `piutang.conf` tidak cocok dengan nama kolom aktual di `Master_temp.xlsx`. Buka file master dan periksa nama kolom persis.

### ❌ `2_ComparePending: Kolom 'No. Faktur' tidak ditemukan di Piutang_clean_temp.xlsx`
Skrip 1 mungkin gagal menghasilkan kolom yang benar. Jalankan Skrip 1 secara manual dan periksa isi `Piutang_clean_temp.xlsx`.

### ❌ `PermissionError` / Excel tidak bisa membuka TEMPLATE.xlsm
`TEMPLATE.xlsm` sedang terbuka di Excel. Tutup semua file Excel, lalu jalankan ulang.

### ❌ Skrip 3 gagal karena tidak ada Microsoft Excel
Aktifkan mode Pure Python: ubah `[PURE] pr_process = Ya` di `piutang.conf`. Namun `3_CalculateAR.py` (xlwings) masih dieksekusi lebih dulu oleh orkestrator dan akan gagal. Untuk sistem tanpa Excel, jalankan skrip secara manual mulai dari `3_CalculateARPurePython.py`.

### ❌ Dialog printer tidak muncul / `6_PrintByPython.py` error
Skrip 6 hanya berjalan di **Windows** dan membutuhkan `pywin32`. Pastikan `pip install pywin32` sudah berhasil. Di Linux/Mac, nonaktifkan dengan `[PPRINT] status = No`.

### ❌ `Error: URL Google Spreadsheet tidak ditemukan di piutang.conf`
Seksi `[SS]` belum diisi. Tambahkan `url =` dan `sheet_name =`.

### ❌ Error autentikasi Google
Periksa `credentials.json` — pastikan `private_key` tersalin lengkap termasuk header dan footer PEM.

---

## 📌 Catatan Penting

- **`Piutang.xls` disalin, bukan dipindahkan** — File asli di folder utama tetap aman.
- **`Giro.xls` opsional** — Jika tidak ada, langkah giro diskip tanpa error (selama `giro_stats = No` atau file Giro memang tidak perlu).
- **`credentials.json` bersifat rahasia** — Tambahkan ke `.gitignore`. Jangan commit ke repositori publik.
- **`TEMPLATE.xlsm` wajib ada** — Baris 1–4 = header, baris 5 = baris data, baris 6 = baris total, baris 7+ = footer TTD. Ganti nama depo sesuai identitas cabang.
- **Jika `pr_process = Ya`, Pure Python menimpa xlwings** — Kedua skrip step 3 selalu dieksekusi; yang terakhir (Pure Python) menentukan output final.
- **Skrip 1D selalu berjalan** — Pembersihan saldo nol tidak punya flag; selalu aktif.
- **`Print_AR.xlsm` di folder utama tidak dihapus** — Hanya salinan di `Dapur/` yang dibersihkan di akhir.
- **Jangan ubah struktur `Dapur/`** — Semua skrip bergantung pada nama file sementara yang sudah ditentukan.

---

*Dikembangkan oleh [ACC-TAX-REIGHTEEN](https://github.com/ACC-TAX-REIGHTEEN)*
