# Automation Account Receivable (AR) Processing System

Sistem ini dirancang untuk melakukan otomatisasi pengolahan data piutang (Account Receivable) mulai dari pembersihan data mentah dari sistem akuntansi, pemfilteran berdasarkan sales/kolektor, penyusunan ke dalam template laporan excel, hingga injeksi data secara otomatis ke Google Sheets.

## Deskripsi Singkat
Program ini bekerja secara sekuensial dengan memproses file ekspor (ExportFile.xls) dan mengubahnya menjadi laporan yang siap cetak serta database digital di Google Sheets. Seluruh proses dikendalikan oleh satu skrip utama untuk memudahkan operasional.

## Struktur File
Berikut adalah struktur folder yang diperlukan agar program berjalan dengan lancar:

- Ambil AR.py (Skrip utama untuk menjalankan seluruh proses)
- ExportFile.xls (Data mentah input)
- TEMPLATE.xlsx (Template laporan untuk output excel)
- Dapur/ (Folder kerja sistem)
    - 1_CleanerAcc.py
    - 2_FilterAR.py
    - 3_CalculateAR.py
    - 4_HelperCleaningData.py
    - 5_InjectDataToSS.py
    - piutang.conf (File konfigurasi sales dan metadata)
    - credentials.json (Kredensial Google Service Account)

## Persyaratan Sistem
Pastikan Python sudah terinstal di sistem Anda beserta pustaka (library) berikut:
- pandas
- numpy
- openpyxl
- xlsxwriter
- gspread
- google-auth

Instalasi library dapat dilakukan dengan perintah:
pip install pandas numpy openpyxl xlsxwriter gspread google-auth

## Konfigurasi

### 1. Konfigurasi piutang.conf
File ini digunakan untuk memetakan Kode Pelanggan ke Nama Sales tertentu serta mengatur metadata laporan.
Format pengisian:
- [NAMA SALES]: Diikuti nama sales.
- [KODE PELANGGAN]: Diikuti daftar kode pelanggan di bawahnya.
- [PERUSAHAAN], [DIVISI], [TANGGAL], [INPUT]: Metadata untuk keperluan upload ke Google Sheets.

### 2. Google Sheets API
- Letakkan file 'credentials.json' yang didapat dari Google Cloud Console ke dalam folder 'Dapur'.
- Pastikan ID Spreadsheet dan ID Worksheet sudah dikonfigurasi pada file '5_InjectDataToSS.py'.
- Berikan akses 'Editor' pada email Service Account Anda di file Google Sheets tujuan.

## Alur Kerja Program

Program dijalankan secara berurutan melalui 'Ambil AR.py' dengan tahapan sebagai berikut:

1. 1_CleanerAcc.py: Membaca file 'ExportFile.xls', mencari header yang sesuai secara dinamis, membersihkan spasi, dan memformat kolom nilai ke dalam tipe data float.
2. 2_FilterAR.py: Memfilter data hasil pembersihan berdasarkan daftar pelanggan yang ada di 'piutang.conf'. Menghitung nilai 'Terbayar' (Nilai Faktur dikurangi Sisa Piutang).
3. 3_CalculateAR.py: Memasukkan data yang telah difilter ke dalam 'TEMPLATE.xlsx'. Melakukan auto-fit kolom, pewarnaan baris total, dan penyusunan tata letak laporan.
4. 4_HelperCleaningData.py: Melakukan pembersihan tambahan pada file siap cetak untuk menghilangkan baris-baris header/footer yang tidak diperlukan sebelum dikirim ke database.
5. 5_InjectDataToSS.py: Mengambil data bersih akhir dan menyisipkannya (append) ke baris terakhir pada Google Sheets yang telah ditentukan.

## Cara Penggunaan
1. Siapkan file 'ExportFile.xls' di direktori utama (sejajar dengan Ambil AR.py).
2. Pastikan file 'piutang.conf' sudah diperbarui sesuai dengan rute penagihan.
3. Jalankan perintah:
   python "Ambil AR.py"
4. Tunggu hingga proses selesai. Program akan menghapus file temporer secara otomatis dan menyisakan file 'Print_AR.xlsx' sebagai laporan final.

## Catatan Penting
- Jangan mengubah struktur kolom pada 'TEMPLATE.xlsx' karena akan memengaruhi koordinasi pemetaan data pada skrip ke-3.
- Pastikan file 'ExportFile.xls' tidak sedang dibuka oleh aplikasi lain saat program dijalankan.
