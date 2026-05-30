# ITBisa Shop Report Bot

Standalone tool untuk generate laporan analisa penjualan ITBisa Shop dari data Shopee, Tokopedia, Tiktok Shop, Bukalapak (legacy), dan CoD. Dirancang untuk dirun berulang tanpa Claude/internet.

## Setup

```bash
cd itbisa-shop-report-bot
pip install -r requirements.txt
```

Persyaratan: Python 3.10+

## Cara Pakai

### 1. Siapkan data

Letakkan semua file Excel di folder `data/`. Script akan otomatis baca semua file matching pola:
- `Stok_*.xlsx` — file pembelian
- `Jual_*.xlsx` — file penjualan

Contoh struktur:
```
data/
├── Stok_2018_2025.xlsx     # historical stok
├── Stok_2026.xlsx          # current year stok
├── Jual_2018_2025.xlsx     # historical jual
└── Jual_2026.xlsx          # current year jual
```

**Stok file**: harus punya sheet `BisaStok` dengan headers di baris ke-2. File terbaru (mis. `Stok_2026.xlsx`) juga harus punya `BisaHilang` dan `BisaPindahBarang` untuk rekonsiliasi sisa stok (lihat "Sisa Stok" di bawah).

**Jual file**: harus punya minimal sheet `BisaJualShopee`. Optional sheets yang akan di-load otomatis kalau ada:
- `BisaJualTiktok` (berisi Tiktok + Tokopedia setelah merger April 2025)
- `BisaJualTokopedia` (legacy, sebelum merger)
- `BisaJualBukalapak` (legacy, platform sudah tutup)
- `BisaJualCoD` (offline/Cash on Delivery)

### 2. Jalankan analisa

```bash
python main.py                  # default: analisa sales tahun berjalan
python main.py --sales 2026     # analisa sales tahun spesifik
python main.py --sales 2024     # analisa sales tahun historis
python main.py --sales          # analisa sales SEMUA tahun yang ditemukan
python main.py --reorder        # analisa reorder standalone (cepat)
python main.py --ab-test        # analisa A/B test (perubahan harga)
python main.py --all            # RUN SEMUANYA: sales all + reorder + ab-test
python main.py --sales 2024 --data-dir /custom/path --output-dir /custom/out
```

CLI ringkasan:

| Flag | Fungsi |
|---|---|
| (tanpa flag) | Sales tahun berjalan |
| `--sales` | Sales SEMUA tahun (1 file per tahun di output/) |
| `--sales 2026` | Sales tahun spesifik |
| `--reorder` | Reorder standalone (`Analisa_Reorder.xlsx`) |
| `--ab-test` | A/B test (`Analisa_AB_Test.xlsx`). Auto-create template kalau belum ada |
| `--all` | Run semuanya: `--sales` (semua tahun) + `--reorder` + `--ab-test` (kalau template ada) |

Year filter berdasarkan `Tanggal Pesan` (hanya untuk angka penjualan). **Sisa stok TIDAK difilter tahun** — selalu dari workbook periode berjalan (lihat "Sisa Stok").

### 3. Lihat hasil

`output/Analisa_Penjualan_ITBisa_<tahun>.xlsx` dengan 12 sheet:

| Sheet | Isi |
|---|---|
| `00_Summary` | Ringkasan performa + temuan utama (data-driven narrative) |
| `01_Paling_Diminati` | Top 40 SKU by Qty Terjual |
| `02_Profit_Tertinggi` | Top 40 SKU by Total Profit |
| `03_Barang_Rugi` | SKU rugi + rekomendasi harga koreksi |
| `04_Margin_Borderline` | SKU markup 0-30% atas HPP (di bawah floor, wajib review) |
| `05_Kandidat_Naik_Harga` | SKU rekomendasi naik harga + skenario +10%/+15%/+20% |
| `06_Per_Platform` | Breakdown Shopee/Tiktok/Tokopedia/Bukalapak/CoD + top SKU per platform |
| `07_Data_Lengkap_per_SKU` | Full per-SKU table untuk drill-down |
| `08_Supplier_Analysis` | China direct (Ocistok/Martkita) vs Market buy (Shopee/Tokopedia) |
| `09_Reorder_Analysis` | Snapshot ROP per SKU, action buckets (STOCKOUT/URGENT/Now/Soon/Overstock) |
| `10_Reorder_Data_Lengkap` | Full per-SKU reorder data |
| `11_Rekap_Stok_per_Gudang` | **BARU**: rekonsiliasi sisa stok per SKU per gudang (reproduksi BisaRekapBarang) |

Reorder data juga tersedia sebagai laporan mandiri via `python main.py --reorder` → `Analisa_Reorder.xlsx` (sheet `03_Rekap_Stok_per_Gudang` ikut disertakan).

## Sisa Stok — Rekonsiliasi ke BisaRekapBarang (PENTING)

Sisa stok bot **dihitung ulang dari workbook periode berjalan** dengan rumus yang sama persis dengan rekap Google Sheets (`BisaRekapBarang`), bukan dari akumulasi lintas tahun. Tujuannya: `sisa_stok` bot = `Total Stok` di BisaRekapBarang.

**Workbook periode berjalan** = file **`Stok_*.xlsx` dan `Jual_*.xlsx` paling baru menurut nama** (mis. `Stok_2026.xlsx`, `Jual_2026.xlsx` — sortir nama, ambil terakhir). Migrasi di file ini = saldo awal periode.

**Rumus per (SKU, gudang):**
```
sisa = Σ beli(sudah sampai) − Σ jual(non-void) + Σ ketemu − Σ hilang
       + Σ pindah_masuk − Σ pindah_keluar
```
- **beli "sudah sampai"** = baris `BisaStok` yang kolom `Tanggal Sampai`-nya **terisi** (kosong = masih di jalan, belum dihitung). Migrasi termasuk (saldo awal).
- **jual non-void** = semua sheet `BisaJual*` di file jual terbaru (termasuk Blibli/Investasi, sesuai cakupan rekap), kecuali `Void = True`.
- **ketemu/hilang** = sheet `BisaHilang` (kolom `Banyak Ketemu` / `Banyak Hilang`) di file stok terbaru.
- **pindah** = sheet `BisaPindahBarang`: `+` ke `Lokasi Penambahan Barang`, `−` dari `Lokasi Pengurangan Barang`. Net nol di total SKU, hanya menggeser antar gudang.

**Total per SKU** = jumlah semua gudang. Inilah `sisa_stok` yang dipakai sheet 07 dan analisa reorder. SKU yang **tidak ada di workbook berjalan** (tidak pernah dibeli di periode ini) → sisa 0 (sama seperti rekap, yang hanya melacak SKU ber-pembelian).

**Saldo gudang negatif** (jual/pindah salah tag ke gudang yang tidak punya stok) di-floor ke 0 dan defisitnya digeser ke gudang yang benar — total tetap utuh, cocok dengan cara rekap menampilkannya. SKU yang **OVERSOLD** (total stok negatif = beneran kurang catat) dibiarkan negatif dan **di-flag ke console** untuk dicek manual.

**Velocity reorder tetap pakai jual lintas tahun** (semua file) — hanya `sisa_stok` yang dari workbook berjalan. HPP juga tetap dari semua file (lihat HPP di bawah).

### Normalisasi SKU (case-insensitive)
Semua SKU di-normalisasi `UPPER().strip()` saat load (stok, jual, hilang, pindah). Ini menyamakan varian beda-kapital (mis. `...PCB-5X7...` vs `...5x7...`) seperti perilaku `SUMIF` Google Sheets yang case-insensitive. Tanpa ini, pandas (case-sensitive) memecah SKU yang sama jadi dua dan salah hitung.

### Dedup Stok — DIHAPUS
Dedup lama `drop_duplicates(SKU, Tanggal Bayar, Qty, Total HPP)` **sudah dihapus**. Dedup itu keliru membuang lot pembelian **asli** yang kebetulan kembar persis (mis. 3× `JUMPER-MF-20CM` di timestamp/qty/harga sama = 6000 pcs, tersisa cuma 2000) — total 14.258 pcs di 15 SKU hilang, dan bobot HPP ikut terdistorsi. Rekap Google Sheets tidak pernah dedup. Sekarang bot hanya menjalankan **drop-Migrasi** (buang baris Migrasi kalau SKU-nya sudah punya pembelian asli).

## Reorder Analysis

Untuk menjawab: **"Kapan saya harus restock SKU ini, dan berapa banyak?"**

Output: sheet `09_Reorder_Analysis` (dalam yearly report) atau file mandiri `Analisa_Reorder.xlsx` (mode `--reorder`). Action buckets urut prioritas: 🔴 STOCKOUT, 🔴 URGENT, 🟠 Now, 🟡 Soon, 🔵 Overstock; plus 💤 Slow/Dead (velocity < 0.5/bulan, skip reorder rule).

**Sisa stok reorder** = on-hand dari ledger workbook berjalan (lihat "Sisa Stok" di atas), **bukan** akumulasi lintas tahun. Karena itu angka reorder kini mencerminkan stok fisik aktual (sudah memperhitungkan BisaHilang & BisaPindahBarang), dan berbeda dari versi lama yang memakai `total beli all-time − total jual all-time`.

### Metodologi (ringkas)

**Velocity**: rata-rata qty terjual/bulan. Window 6 bulan default; fallback 12mo lalu 24mo bila < 3 bulan aktif. Bulan tanpa transaksi = 0.

**Volatility (CV)** = std/mean qty bulanan → Stabil (<0.3, safety 1.2×), Moderate (<0.7, 1.5×), Volatile (≥0.7, 2.0×).

**Lead time**: China direct (≥50% qty dari China) = 2 bulan; Market buy = 0.25 bulan.

**ROP** = MAX dari (a) `lead_demand × safety_multiplier`, (b) `lead_demand + max_single_order_12mo` (proteksi bulk buyer). `lead_demand = velocity × lead_time`.

**Suggested order** = `velocity × 6 − sisa + lead_demand`.

### Konfigurasi Reorder

Edit `config.py`:

| Parameter | Default | Arti |
|---|---|---|
| `LEAD_TIME_CHINA_MONTHS` | `2.0` | Lead time order China direct |
| `LEAD_TIME_MARKET_MONTHS` | `0.25` | Lead time market buy (≈ 1 minggu) |
| `BULK_CHINA_SHARE_THRESHOLD` | `0.5` | Persen qty dari China untuk klasifikasi bulk-China |
| `VELOCITY_MIN_ACTIVE_MONTHS` | `3` | Min bulan aktif di window 6mo sebelum fallback ke 12mo |
| `CV_STABLE_MAX` / `CV_MODERATE_MAX` | `0.3` / `0.7` | Threshold volatility |
| `SAFETY_MULT_STABLE/MODERATE/VOLATILE` | `1.2` / `1.5` / `2.0` | Safety multiplier per kategori |
| `TARGET_MONTHS_POST_REORDER` | `6` | Bulan cadangan setelah reorder masuk |
| `SLOW_DEAD_MAX_VELOCITY` | `0.5` | Velocity di bawah ini = Slow/Dead, skip reorder rule |
| `ROP_URGENT/NOW/SOON_RATIO` | `0.7` / `1.0` / `1.3` | Rasio sisa/ROP untuk threshold status |
| `OVERSTOCK_MONTHS` | `12.0` | Bulan cover di atas ini = Overstock |

## Sheet 08: Supplier Analysis

4 sub-section: A) China vs Market (flag merah kalau Market lebih murah), B) China-only HPP tidak konsisten (CV > 15%), C) Top China-only, D) Top Market-only.

Klasifikasi: **China** = `Luar Negeri? = 1` ATAU nama toko mengandung Ocistok/Martkita/Aliexpress/Jasa Impor/1688/Alibaba. **Market** = Shopee/Tokopedia/Bukalapak/Tiktok. **Other** = selain itu.

## Struktur Project

```
itbisa-shop-report-bot/
├── README.md
├── requirements.txt
├── .gitignore
├── config.py             # constants: thresholds, colors, column names (stok/jual/hilang/pindah), supplier keywords, reorder + ab-test params
├── data_loader.py        # multi-file glob loading + SKU normalisasi + loader workbook berjalan (stok arrived/jual/hilang/pindah)
├── analysis.py           # HPP WA, profit, per-SKU agg, supplier classify, reorder metrics, build_stock_ledger
├── tables.py             # table builders (diminati, profit, rugi, kandidat, supplier, reorder)
├── excel_writer.py       # Excel output + sheet rekap stok per gudang
├── ab_testing.py         # A/B test analyzer
├── main.py               # CLI entry point
├── data/                 # input files (gitignored)
└── output/               # generated reports
```

## Metodologi HPP & Pricing

### HPP (Harga Pokok Penjualan) — Weighted Average dengan Ocistok-Priority
- Kalau SKU ada pembelian Ocistok/Martkita (China direct) → WA dari pembelian Ocistok/Martkita **saja**.
- Kalau tidak → WA dari semua pembelian.

Combined dari semua `Stok_*.xlsx`. Karena dedup agresif sudah dihapus, bobot HPP kini memakai **semua lot pembelian asli** (lebih akurat). Kolom `HPP Source` di sheet 07 menunjukkan source per SKU.

Inventory display `total_qty_beli` tetap pakai semua pembelian (after drop-Migrasi). `sisa_stok` TIDAK lagi `total_qty_beli − jual` — sekarang dari ledger workbook berjalan.

### Markup Floor 30% atas HPP (UNIFORM)
Harga jual minimum = `HPP × 1.30`. **Markup** = `(Jual − HPP)/HPP`; **Margin** = `Profit/Omzet` (after admin). Sheet 07 menampilkan keduanya. Pricing aktual tetap competitor-aware; analyzer hanya nge-flag.

### Profit per Transaksi
```
Untung = Omzet − (HPP_WA × Qty_Jual) + Tambahan + Kode_Unik
```
`Tambahan` & `Kode_Unik` sudah signed (negatif = biaya admin), tidak di-negate.

### Exclusions Otomatis (data jual)
SKU null, Invoice `"Dummy"`, `Void = True`, `EXCLUDED_SKUS`, baris invalid numeric, lalu filter tahun (kecuali reorder/ab-test). SKU dijual tanpa HPP → di-warning, exclude dari profit.

### Migrasi Entries (Inventory Carry-Over)
Akhir tahun, sisa stok dicatat ke file tahun baru dengan prefix `"Migrasi - "` di kolom Toko. Untuk **sisa stok**, Migrasi = saldo awal periode (dihitung penuh di ledger). Untuk **HPP/total beli lintas tahun**, Migrasi di-drop kalau SKU sudah punya pembelian asli (hindari double-count).

## A/B Testing — Track Perubahan Harga

Config `data/ab_tests.xlsx` (sheet `BisaABTest`): kolom `SKU`, `Tanggal Perubahan`, `Nama Test`, `Catatan`. Template auto-create saat `--ab-test` pertama. Output `output/Analisa_AB_Test.xlsx`: pre vs post change date, daily-rate normalized, verdict ✅/🟡/🔴/⚪.

## Catatan

- **Standalone**, **idempotent**, ditest dengan 15.000+ baris transaksi.
- **Rekonsiliasi sisa stok**: sudah diverifikasi cocok dengan `BisaRekapBarang` (total per SKU exact). Pencocokan mensyaratkan SKU di sumber sudah konsisten kapital (case sudah dinormalisasi otomatis oleh bot).
