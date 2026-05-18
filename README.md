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

**Stok file**: harus punya sheet `BisaStok` dengan headers di baris ke-2.

**Jual file**: harus punya minimal sheet `BisaJualShopee`. Optional sheets yang akan di-load otomatis kalau ada:
- `BisaJualTiktok` (berisi Tiktok + Tokopedia setelah merger April 2025)
- `BisaJualTokopedia` (legacy, sebelum merger)
- `BisaJualBukalapak` (legacy, platform sudah tutup)
- `BisaJualCoD` (offline/Cash on Delivery)

### 2. Jalankan analisa

```bash
python main.py                  # analisa tahun berjalan
python main.py --year 2026      # analisa tahun spesifik
python main.py --year 2024      # analisa tahun historis
python main.py --all            # SEMUA tahun yang ditemukan di data (sekali run)
python main.py --reorder        # analisa reorder standalone (cepat)
python main.py --ab-test        # analisa A/B test (perubahan harga)
python main.py --year 2024 --data-dir /custom/path --output-dir /custom/out
```

Year filter berdasarkan `Tanggal Pesan`. Semua file jual di folder di-load, lalu di-filter ke tahun yang diminta.

Mode `--all` generate satu file Excel per tahun (mis. `Analisa_Penjualan_ITBisa_2018.xlsx` sampai `..._2026.xlsx`) dan tampilkan ringkasan profit semua tahun di akhir console output. Cocok untuk lihat tren multi-tahun.

Mode `--reorder` generate `output/Analisa_Reorder.xlsx` yang fokus ke pertanyaan "kapan harus restock dan berapa banyak?". Output mandiri tanpa analisa tahunan — lebih cepat. Berlaku snapshot per tanggal run (forward-looking, bukan retrospective). Lihat section "Reorder Analysis" di bawah.

Mode `--ab-test` baca config dari `data/ab_tests.xlsx` (template di-auto-create kalau belum ada), generate `output/Analisa_AB_Test.xlsx` dengan perbandingan metrik sebelum vs sesudah perubahan harga. Lihat section "A/B Testing" di bawah.

### 3. Lihat hasil

`output/Analisa_Penjualan_ITBisa_<tahun>.xlsx` dengan 11 sheet:

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
| `09_Reorder_Analysis` | **BARU**: snapshot ROP per SKU, action buckets (STOCKOUT/URGENT/Now/Soon/Overstock) |
| `10_Reorder_Data_Lengkap` | **BARU**: full per-SKU reorder data |

Reorder data juga tersedia sebagai laporan mandiri via `python main.py --reorder` → `Analisa_Reorder.xlsx`.

## Reorder Analysis

Untuk menjawab pertanyaan: **"Kapan saya harus restock SKU ini, dan berapa banyak?"**

### Output

Sheet `09_Reorder_Analysis` (dalam yearly report) atau file mandiri `Analisa_Reorder.xlsx` (mode `--reorder`). Strukturnya:

- **Metodologi**: ringkasan rumus dan parameter
- **Ringkasan Status**: jumlah SKU per kategori
- **Action buckets** (urut prioritas):
    - A. 🔴 **STOCKOUT** — sisa ≤ 0, customer sudah tidak bisa beli
    - B. 🔴 **Reorder URGENT** — sisa < ROP × 0.7, ada resiko stockout sebelum reorder masuk
    - C. 🟠 **Reorder Now** — sisa < ROP, reorder minggu ini
    - D. 🟡 **Reorder Soon** — sisa < ROP × 1.3, mulai siap-siap
    - E. 🔵 **Overstock** — sisa > 12 bulan cadangan, stop reorder

SKU dengan velocity < 0.5/bulan diklasifikasi `💤 Slow/Dead` dan tidak masuk reorder rule (cuma muncul di data lengkap).

### Metodologi

**Velocity (kecepatan jual per SKU)**: rata-rata qty terjual per bulan. Default window 6 bulan terakhir. Kalau window 6 bulan punya < 3 bulan aktif, fallback ke 12 bulan; kalau 12 bulan juga tipis, fallback 24 bulan. Bulan tanpa transaksi dihitung 0 (penting biar SKU yang cuma laku 1 bulan tidak over-stated).

**Volatility (CV - coefficient of variation)**: std deviation / mean dari qty bulanan. Mengukur seberapa "spiky" demand-nya.
- `CV < 0.3` → **Stabil** (safety multiplier 1.2×)
- `0.3 ≤ CV < 0.7` → **Moderate** (safety 1.5×)
- `CV ≥ 0.7` → **Volatile** (safety 2.0×)

**Lead time**: berapa lama dari order ke barang siap dijual.
- **China direct** (Ocistok/Martkita/Aliexpress): default 2 bulan
- **Market buy** (Shopee/Tokopedia/Bukalapak/Tiktok): default 0.25 bulan (≈ 1 minggu)

SKU diklasifikasi "bulk China" kalau ≥ 50% qty historisnya dari sumber China. Pembelian market kecil sebagai gap-fill tidak menggeser klasifikasi.

**ROP (Reorder Point)** = MAX dari dua proteksi:
1. **Safety-based** = `lead_demand × safety_multiplier` (lead_demand = velocity × lead_time)
2. **Bulk-order protection** = `lead_demand + max_single_order_last_12mo` — biar satu pembeli grosir besar tidak langsung bikin stockout

Pakai yang lebih tinggi.

**Suggested order quantity** = `target_qty_post_reorder − sisa + lead_demand` di mana `target_qty = velocity × 6 bulan`. Artinya: pesan cukup banyak supaya setelah barang datang, stok cukup buat 6 bulan ke depan.

**Sisa stok** = `total qty beli all-time − total qty jual all-time`. Bukan hanya tahun analisa, biar akurat (sisa stok lintas tahun terbawa).

### Contoh: NE555P (kondisi per 18 Mei 2026)

| Metric | Nilai |
|---|---|
| Sisa stok | 6,558 |
| Velocity 6mo | 1,527/bulan |
| CV | 0.99 → Volatile |
| Max 1 order (12mo) | 2,000 pcs |
| Lead time | 2 bulan (78% qty dari China) |
| Lead demand | 3,053 |
| ROP safety | 6,106 (= 3,053 × 2.0) |
| ROP bulk | 5,053 (= 3,053 + 2,000) |
| **ROP final** | **6,106** ← reorder kalau sisa di bawah ini |
| Status | 🟡 Reorder Soon |

Aturan lama "reorder kalau < 3,000" = sama dengan lead demand tanpa safety buffer apapun. Risiko: satu pembeli grosir besar selama lead time bisa langsung bikin stockout.

### Konfigurasi Reorder

Edit `config.py` untuk adjust default:

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

4 sub-section untuk analisa pengadaan stok:

- **A. Perbandingan China vs Market** — SKU yang punya pembelian dari KEDUA sumber. Flag merah kalau Market lebih murah dari China = pertimbangkan stop reorder China.
- **B. China-only dengan HPP tidak konsisten** — SKU yang HPP impor variansinya tinggi (CV > 15%). Sinyal kurs goyang atau supplier ganti spec.
- **C. Top China-only** — SKU yang cuma diimpor, belum pernah test market. Potensi diversifikasi sumber.
- **D. Top Market-only** — SKU yang cuma beli dari market, belum pernah test import. Potensi reduksi biaya.

### Klasifikasi Supplier
- **China**: `Luar Negeri? = 1` ATAU nama toko mengandung: Ocistok, Martkita, Aliexpress, "Jasa Impor", 1688, Alibaba
- **Market**: nama toko mengandung: Shopee, Tokopedia, Bukalapak, Tiktok
- **Other**: selain itu (URO Shop, PT, dll.)

## Struktur Project

```
itbisa-shop-report-bot/
├── README.md
├── requirements.txt
├── .gitignore
├── config.py             # constants: thresholds, colors, column names, supplier keywords, reorder params
├── data_loader.py        # multi-file glob loading with optional sheets
├── analysis.py           # HPP weighted avg, profit, per-SKU aggregation, supplier classify, reorder metrics
├── tables.py             # table builders (diminati, profit, rugi, kandidat, supplier, reorder, dll.)
├── excel_writer.py       # Excel output with styling helpers, reorder sheet writers
├── ab_testing.py         # A/B test analyzer
├── main.py               # CLI entry point
├── data/                 # input files (gitignored)
└── output/               # generated reports
```

## Metodologi

### HPP (Harga Pokok Penjualan)
**Weighted Average dengan Ocistok-Priority**:
- Kalau SKU ada pembelian dari Ocistok/Martkita (China direct) → WA dari pembelian Ocistok/Martkita **saja** (abaikan market buy & supplier lain)
- Kalau tidak → WA dari pembelian yang ada (semua supplier)

Rasionalnya: Ocistok/Martkita adalah supplier utama dan stabil. HPP dari market-buy biasanya lebih mahal/fluktuatif (untuk testing pasar). Untuk cost basis yang representatif ke kondisi sourcing aktual, prioritaskan Ocistok.

Inventory metrics (`total_qty_beli`, `sisa_stok`) tetap pakai SEMUA pembelian (tidak filtered ke Ocistok saja).

Combined dari semua `Stok_*.xlsx` files di folder.
```
HPP_WA[sku] = sum(Total_HPP_Rp) / sum(Banyak_Barang_Buah)
  ↳ filtered ke pembelian Ocistok/Martkita kalau ada, else pakai semua
```

Kolom `HPP Source` di sheet `07_Data_Lengkap_per_SKU` menunjukkan source-nya per SKU.

### Margin Floor: Markup 30% UNIFORM
**Semua SKU wajib markup minimum 30% atas HPP** untuk dianggap "aman tidak rugi".

Rule: **Harga jual minimum = HPP × 1.30**

Contoh: HPP Rp 1,000 → jual minimum Rp 1,300. Setelah dipotong biaya admin Shopee (~12%) atau Tiktok (~21%), masih ada untung tipis atau balik modal.

**Bedakan dengan gross margin** (yang sering disebut "margin" umumnya):
- **Markup** (yang dipakai di sini): `(Jual − HPP) / HPP × 100` — over cost
- **Margin** (kontekstual): `Profit / Omzet × 100` — after admin/biaya

Threshold ini berlaku untuk:
- Sheet `04_Margin_Borderline` → SKU dengan markup 0-30% (di bawah floor, wajib review)
- Sheet `05_Kandidat_Naik_Harga` → filter markup ≥ 30% (sudah di atas floor, layak push lebih tinggi)
- Sheet `03_Barang_Rugi` → rekomendasi harga koreksi = HPP × 1.30

Sheet `07_Data_Lengkap_per_SKU` menampilkan **dua kolom**: Markup % (rule) dan Margin % (konteks).

Catatan: pricing aktual tetap competitor-aware. Kalau kompetitor harga lebih rendah, Anda boleh terima markup di bawah 30% (mis. untuk SKU stok lama harus dihabiskan) — tapi laporan tetap flag sebagai "perlu attention".

### Profit per Transaksi
```
Untung = Omzet − (HPP_WA × Qty_Jual) + Tambahan + Kode_Unik
```
`Tambahan` dan `Kode_Unik` sudah signed (negatif = biaya admin). Tidak di-negate lagi.

### Exclusions Otomatis (data jual)
1. Baris dengan SKU null
2. Invoice diawali `"Dummy"` (test data)
3. `Void = True`
4. SKU di `EXCLUDED_SKUS` (default: `ITBISA-BUBBLE-WRAP`)
5. Baris invalid numeric (header row scatter)
6. Filter tahun analisa berdasarkan `Tanggal Pesan` (kecuali `--reorder` yang lintas tahun)

SKU yang dijual tanpa HPP: di-warning ke console, di-exclude dari profit.

### Dedup Stok
Saat load multiple stok files, dedup berdasarkan `(SKU, Tanggal_Bayar, Qty_Beli, Total_HPP)` — keep first.

### Migrasi Entries (Inventory Carry-Over)
Di akhir tahun (31 Des 23:59), file `Stok_<tahun-1>.xlsx` "ditutup" dan **sisa stok yang belum terjual** dicatat ulang ke file tahun baru dengan prefix `"Migrasi - "` di kolom Toko. Ini cuma penanda referensi inventory, bukan pembelian baru — supaya buku kas tidak double-counted.

**Masalah jika di-treat sebagai pembelian normal:**
- HPP_WA terinflasi/terdistorsi (cost basis dihitung 2x: dari historical asli + dari Migrasi)
- Total qty beli over-counted (Migrasi qty + original qty)
- Sisa stok over-counted

**Penanganan otomatis di script:**
- Migrasi entries (deteksi via prefix `MIGRASI_PREFIX` di `config.py`) **di-drop dari HPP_WA dan total_qty_beli** kalau SKU yang sama juga punya non-Migrasi data (dari file historis atau pembelian baru)
- Migrasi entries **dipertahankan** kalau itu satu-satunya sumber HPP untuk SKU tersebut (mis., user cuma load Stok_2026 tanpa historis) — supaya tidak kehilangan data
- Sisa stok dihitung pakai **all-time qty terjual** (bukan cuma tahun analisa), supaya sisa stok = total real bought - total real sold

Output console contoh:
```
→ Migrasi: drop 125 duplikat (SKU sudah ada di non-Migrasi), keep 0 (satu-satunya sumber HPP)
```

### Kandidat Naik Harga
SKU yang lolos 3 filter:
- Qty terjual ≥ persentil 80 (top 20% paling laku)
- Markup ≥ 30% atas HPP (di atas floor)
- Sisa stok > 0

Scoring:
```
Score = 0.6 × normalize(Qty_Terjual) + 0.4 × normalize(Markup_pct)
```

## Konfigurasi

Edit `config.py`:

| Parameter | Default | Arti |
|---|---|---|
| `QTY_PERCENTILE` | `0.80` | Persentil qty untuk kandidat |
| `MARKUP_THRESHOLD_KANDIDAT` | `30.0` | Markup minimum (%) atas HPP untuk kandidat |
| `MARKUP_BORDERLINE_MIN/MAX` | `0.0 / 30.0` | Range markup borderline (%) |
| `TARGET_MARKUP_KOREKSI` | `0.30` | Target markup atas HPP untuk rekomendasi harga koreksi (jual = HPP × 1.30) |
| `PRICE_SCENARIOS` | `[0.10, 0.15, 0.20]` | Skenario kenaikan harga |
| `CHINA_KEYWORDS` | `["ocistok", "martkita", ...]` | Keyword untuk klasifikasi supplier China |
| `MARKET_KEYWORDS` | `["shopee ", ...]` | Keyword untuk klasifikasi supplier Market |
| `HPP_VARIANCE_THRESHOLD` | `0.15` | CV threshold (15%) untuk flag HPP volatile |
| `EXCLUDED_SKUS` | `{"ITBISA-BUBBLE-WRAP"}` | SKU yang di-exclude |

Reorder-specific parameters di section terpisah — lihat "Konfigurasi Reorder" di atas.

## Console Output

```
============================================================
ANALISA PENJUALAN ITBISA — TAHUN 2026
============================================================

✓ Membaca stok: Stok_2018_2025.xlsx
✓ Membaca stok: Stok_2026.xlsx
  → Stok bersih: 1,168 baris (dari 1,285 mentah, hapus 23 duplikat)
✓ Membaca jual: Jual_2018_2025.xlsx
  → BisaJualShopee: 7,363 baris
  → BisaJualTiktok: 692 baris
  → BisaJualTokopedia: 3,763 baris
  → BisaJualBukalapak: 359 baris
  → BisaJualCoD: 585 baris
✓ Membaca jual: Jual_2026.xlsx
✓ Cleaning jual:
    Mentah               :  15,467
    ...
    Filter tahun 2026    : buang 12,321
    Bersih               :   2,660
✓ HPP weighted average: 177 SKU
✓ Reorder analysis: 188 SKU (STOCKOUT: 9, URGENT: 12, Now: 3, Soon: 2, Healthy: 23, Overstock: 76, Dead: 63)
✓ Aggregasi per SKU
✓ Membangun tabel analisa
  Qty threshold (p80): 806 pcs
✓ Menulis laporan ke output/Analisa_Penjualan_ITBisa_2026.xlsx
```

## Workflow Bulanan

1. Update `Jual_2026.xlsx` dari export Shopee/Tiktok terbaru
2. Update `Stok_2026.xlsx` dengan pembelian baru
3. Run `python main.py` (default current year) → laporan lengkap dengan reorder sheet
4. Atau quick check: `python main.py --reorder` → cuma laporan reorder
5. Buka Excel output, baca sheet `00_Summary` dulu untuk overview cepat, lalu `09_Reorder_Analysis` untuk action list

## Troubleshooting

**`Tidak ada file Stok ditemukan`** — pastikan ada minimal 1 file `Stok_*.xlsx` di folder `data/`.

**`Kolom hilang`** — struktur Excel berubah. Cek nama kolom raw di Excel vs `config.py`.

**`Tidak ada data jual untuk tahun X`** — tahun X tidak punya transaksi setelah filtering. Cek `Tanggal Pesan` di file jual.

**Pandas warning soal data validation** — bisa diabaikan, limitasi openpyxl saat read xlsx dengan dropdown/data validation.

## A/B Testing — Track Perubahan Harga

Untuk menjawab pertanyaan: **"Setelah saya naikkan harga X, apakah penjualan turun atau tetap stabil?"**

### Workflow

1. **Pertama kali**: jalankan `python main.py --ab-test`. Script otomatis bikin template `data/ab_tests.xlsx`.

2. **Isi template**: tambahkan baris untuk setiap perubahan harga yang mau dilacak:

| SKU | Tanggal Perubahan | Nama Test | Catatan |
|---|---|---|---|
| ITBISA-IC-NE555P-DIP8 | 2026-05-17 | NE555P Price Bump May 2026 | 599 → 699 (1pcs), 50pcs → 689, 1000pcs → 679 |
| ITBISA-IC-PC817-DIP4 | 2026-06-01 | PC817 normalisasi | dari 480 ke 550 |

3. **Run analisa**: `python main.py --ab-test` → output `output/Analisa_AB_Test.xlsx`

### Apa yang Dihitung

Untuk setiap test:
- **Pre-period**: semua transaksi SKU itu SEBELUM tanggal perubahan (full data, bisa multi-tahun)
- **Post-period**: semua transaksi SKU itu SEJAK tanggal perubahan
- **Daily rates** (qty/day, omzet/day, profit/day) untuk fair comparison karena window pre vs post biasanya beda panjang
- **Avg unit price** sebelum vs sesudah → verify perubahan harga benar terdeteksi di data
- **Markup % & margin %** sebelum vs sesudah
- **Delta %** untuk setiap metric
- **Verdict** otomatis: ✅ Effective / 🟡 Mixed / 🔴 Bad / ⚪ Inconclusive

### Verdict Logic

- ✅ **Effective**: profit/day naik > 5%, qty/day turun < 10%
- 🟡 **Mixed**: profit/day naik > 5%, tapi qty/day turun > 10% (margin naik tapi customer berkurang)
- 🔴 **Bad**: profit/day turun > 5%
- ⚪ **Inconclusive**: data sedikit (< 3 hari post) atau perubahan tidak signifikan

### Catatan
- Harga otomatis di-derive dari `Omzet / Qty_Jual` per transaksi — Anda tidak perlu input harga lama/baru manual
- Multi-tier (1pcs/50pcs/1000pcs) di-blend menjadi avg unit price (tidak dipisah per tier)
- Window full data: pre bisa multi-tahun, post mulai dari tanggal perubahan
- Disarankan tunggu minimal 1-2 minggu sebelum interpretasi hasil (default warning di < 3 hari)
- HPP yang dipakai untuk profit calc adalah HPP_WA terkini (dari semua stok files)

## Catatan

- **Standalone**: bisa dirun tanpa Claude/AI
- **Idempotent**: input sama → output sama (kecuali timestamp di Summary)
- **Memory-friendly**: ditest dengan 15,000+ baris transaksi
- **Yearly snapshot**: setiap run `--year`/`--all` filter ke tahun terkait. Mode `--reorder` lintas tahun karena forward-looking.