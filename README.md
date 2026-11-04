# Sedukopi: Analisis Penjualan Kedai Kopi dengan Excel

Proyek portofolio data analyst keempat. Setelah [Python](https://github.com/pajareeeee/gayanara_ecommerce_analysis), [SQL](https://github.com/pajareeeee/gayanara_sql_analysis), dan [Power BI](https://github.com/pajareeeee/sakumaya_powerbi_dashboard), kali ini analisisnya dikerjakan sepenuhnya di **Excel**: validasi data, penggabungan tabel, Pivot Table, dan dashboard satu halaman.

> Dataset simulasi "Sedukopi – Jaringan Kedai Kopi" dari [Ngulik Data](https://ngulikdata.com/datasets/sedukopi). Semua data bersifat fiktif.

[![Dashboard Sedukopi](https://github.com/pajareeeee/sedukopi_excel_analysis/raw/main/images/dashboard.jpeg)](https://github.com/pajareeeee/sedukopi_excel_analysis/blob/main/images/dashboard.jpeg)

## Pertanyaan Bisnis

1. Outlet mana yang revenue-nya tertinggi dan terendah?
2. Bagaimana tren revenue bulanan?
3. Kategori menu mana yang paling besar revenue dan margin-nya?
4. Jam berapa penjualan paling ramai?
5. Tipe order mana (dine in, takeaway, delivery) yang paling besar kontribusinya?
6. Metode pembayaran apa yang menghasilkan pendapatan tertinggi dan terendah?
7. Apakah outlet dengan kursi lebih banyak revenue-nya lebih besar?
8. Bagaimana rasio gaji staf terhadap revenue per outlet?

## Isi Workbook

File: `sedukopi_analisis.xlsx`

| Sheet | Fungsi |
|---|---|
| `Dashboard` | Ringkasan satu halaman: 4 angka utama, 8 grafik, temuan, dan keterbatasan |
| `Catatan` | Tiap pertanyaan bisnis beserta jawaban, catatan, dan rekomendasi |
| `analisis`, `tren`, `kategori`, `jam`, `tipe order`, `metode pembayaran` | Pivot Table dan PivotChart untuk pertanyaan 1-6 |
| `kapasitas`, `staff` | Tabel rumus untuk pertanyaan 7 dan 8 (data dari tabel outlets dan staff) |
| `validasi` | Pengecekan kualitas data dengan rumus |
| `data` | Tabel gabungan satu baris per item terjual, lengkap dengan tanggal, outlet, kategori, harga pokok, margin |
| `raw_*` | Lima tabel asli (orders, details, menu, outlets, staff), tidak diubah |

## Keputusan Data

- Data mentah tidak diubah, tiap CSV diimpor ke sheet `raw_*` dan dijadikan Excel Table
- Validasi dengan rumus: jumlah baris cocok (5.000 order, 8.123 detail), tidak ada `order_id` ganda, tidak ada total kosong, dan total order sama dengan jumlah detail (selisih 0)
- Sheet `data` dibuat dari tabel detail dengan `XLOOKUP` untuk mengambil tanggal, outlet, kategori, harga pokok, tipe order, dan metode bayar, plus kolom hitung biaya, margin, bulan, dan jam
- Revenue = `subtotal` item. Margin = `subtotal` - (`quantity` x `cost_price`), jadi margin kotor
- Analisis staf memakai staf berstatus aktif, karena data hanya mencatat status terakhir
- Perbandingan antar outlet memperhitungkan lama operasi, karena beberapa outlet baru buka pada 2023

## Temuan Utama

- Revenue Rp392,9 juta dari 5.000 order, margin kotor 69,9%, rata-rata Rp78.581 per order
- Revenue tahunan naik dari Rp80 juta (2022) ke Rp140 juta (2024). Bulan tertinggi Mar 2025 (Rp22,1 juta), terendah Apr 2022 (Rp4,8 juta)
- Tujuh jam sibuk (07-08, 12-13, 17-19) menyumbang sekitar 78% revenue. Jam 13 tertinggi, jam 11 paling sepi
- Dine in 50% revenue, takeaway 30%, delivery 20%
- QRIS metode bayar terbesar (25,8%), disusul GoPay (24,0%) dan OVO (19,7%). Kartu Debit terendah (15,0%)
- Makanan kategori terbesar (Rp149 juta), tapi persen margin hampir sama di semua kategori (sekitar 70%)
- Outlet tertinggi OUT003 Senopati (Rp25,5 juta), terendah OUT020 Margonda (Rp16,0 juta)
- Kapasitas kursi tidak terlihat berhubungan dengan revenue (korelasi 0,08)

## Insight dan Rekomendasi

- **Pertumbuhan sebagian karena outlet bertambah.** Outlet aktif naik dari 12 (awal 2022) ke 18 (Juni 2023), jadi revenue yang naik belum tentu berarti tiap outlet makin laris. Lonjakan Jan-Mar 2025 datang dari jumlah order (118-160 per bulan pada Okt-Des 2024 menjadi 231-273 per bulan), bukan dari nilai per order yang tetap di kisaran Rp72-86 ribu. Lonjakan ini terjadi merata di hampir semua dari 18 outlet, bukan karena outlet baru, jadi kemungkinan besar pola dataset simulasi.
- **Jam sibuk menentukan penjadwalan.** Jadwal staf bisa difokuskan ke tiga puncak (pagi, makan siang, sore). Promo di jam sepi layak dicoba, tapi data ini belum bisa menjawab apakah promo akan efektif, jadi perlu diuji dulu.
- **Dine in separuh revenue.** Pengalaman di outlet paling penting. Delivery 20% masih bisa dikembangkan, tapi data belum memuat biaya pengiriman atau komisi platform.
- **Pembayaran digital mendominasi.** QRIS, GoPay, dan OVO menyumbang sekitar 70% revenue. Data belum memuat biaya per metode bayar, jadi belum bisa dinilai metode mana yang paling menguntungkan.
- **Untung ditentukan volume, bukan harga pokok.** Persen margin tiap kategori hampir sama, jadi menaikkan penjualan Makanan dan Coffee lebih berpengaruh daripada mengubah harga pokok.
- **Menambah kursi belum tentu menaikkan penjualan.** Rata-rata revenue per outlet hampir sama di semua kelompok kursi: 50 kursi atau kurang Rp21,8 juta, 55-60 kursi Rp20,9 juta, 65-70 kursi Rp23,4 juta.
- **Revenue per staf perlu dibaca hati-hati.** OUT016 dan OUT009 punya revenue per staf tertinggi, tetapi masing-masing hanya punya 3 staf aktif karena 3 dari 6 stafnya sudah keluar, jadi belum tentu lebih efisien. OUT008 dan OUT010 punya 6 staf aktif dengan revenue per staf rendah. Karena datanya sampel, ini hanya indikasi awal, bukan alasan langsung mengubah jadwal.

## Keterbatasan

- Dataset simulasi dan hanya sampel transaksi (5.000 order, sekitar 8 order per outlet per bulan). Gaji bulanan staf per outlet sekitar 13-44 kali lipat revenue bulanannya, jadi rasio gaji terhadap revenue dan profitabilitas tidak bisa dihitung secara bermakna.
- Margin adalah margin kotor (belum termasuk gaji, sewa, dan biaya operasional)
- Outlet yang buka 2023 (OUT017, OUT019, OUT020) punya periode jualan lebih pendek, jadi perbandingan total antar outlet kurang adil. OUT015 dan OUT018 tidak punya order (status temporarily closed)
- Data staf hanya mencatat status terakhir tanpa tanggal keluar, jadi hanya staf aktif yang dihitung
- 2025 baru tiga bulan (Jan-Mar), jangan dibandingkan dengan total tahun lain

## Pengembangan Berikutnya

- Pivot per hari dalam seminggu (Minggu tertinggi, Rabu terendah, selisih sekitar 18%)
- Pivot per menu: Club Sandwich terbesar, 10 dari 80 menu tidak pernah terjual
- Revenue per staf dari kolom `staff_id` di tabel orders, bukan dari jumlah staf
- Ringkasan per kota dan provinsi

## Cara Membuka

1. Unduh `sedukopi_analisis.xlsx`
2. Buka dengan Microsoft Excel (rumus `XLOOKUP` tidak jalan di versi lama)
3. Mulai dari sheet `Dashboard`, lalu `Catatan` untuk jawaban lengkap tiap pertanyaan

## Skill yang Dipakai

Pivot Table dan PivotChart, Excel Table dan structured reference, `XLOOKUP`, `SUMIFS`, `COUNTIFS`, `DATEDIF`, `CORREL`, Calculated Field, validasi data dengan rumus, desain dashboard, dan penulisan temuan serta keterbatasan analisis.
