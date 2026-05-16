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
python main.py --year 2024 --data-dir /custom/path --output-dir /custom/out
```

Year filter berdasarkan `Tanggal Pesan`. Semua file jual di folder di-load, lalu di-filter ke tahun yang diminta.

### 3. Lihat hasil

`output/Analisa_Penjualan_ITBisa_<tahun>.xlsx` dengan 9 sheet:

| Sheet | Isi |
|---|---|
| `00_Summary` | Ringkasan performa + temuan utama (data-driven narrative) |
| `01_Paling_Diminati` | Top 40 SKU by Qty Terjual |
| `02_Profit_Tertinggi` | Top 40 SKU by Total Profit |
| `03_Barang_Rugi` | SKU rugi + rekomendasi harga koreksi |
| `04_Margin_Borderline` | SKU margin 0-5% (rawan rugi) |
| `05_Kandidat_Naik_Harga` | SKU rekomendasi naik harga + skenario +10%/+15%/+20% |
| `06_Per_Platform` | Breakdown Shopee/Tiktok/Tokopedia/Bukalapak/CoD + top SKU per platform |
| `07_Data_Lengkap_per_SKU` | Full per-SKU table untuk drill-down |
| `08_Supplier_Analysis` | **BARU**: China direct (Ocistok/Martkita) vs Market buy (Shopee/Tokopedia) |

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
├── config.py             # constants: thresholds, colors, column names, supplier keywords
├── data_loader.py        # multi-file glob loading with optional sheets
├── analysis.py           # HPP weighted avg, profit, per-SKU aggregation, supplier classify
├── tables.py             # table builders (diminati, profit, rugi, kandidat, supplier, dll.)
├── excel_writer.py       # Excel output with styling helpers
├── main.py               # CLI entry point
├── data/                 # input files (gitignored)
└── output/               # generated reports
```

## Metodologi

### HPP (Harga Pokok Penjualan)
**Weighted Average** dari SEMUA pembelian historis per SKU. Combined dari semua `Stok_*.xlsx` files di folder.
```
HPP_WA = sum(Total_HPP_Rp) / sum(Banyak_Barang_Buah)  per SKU
```

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
6. Filter tahun analisa berdasarkan `Tanggal Pesan`

SKU yang dijual tanpa HPP: di-warning ke console, di-exclude dari profit.

### Dedup Stok
Saat load multiple stok files, dedup berdasarkan `(SKU, Tanggal_Bayar, Qty_Beli, Total_HPP)` — keep first.

### Kandidat Naik Harga
SKU yang lolos 3 filter:
- Qty terjual ≥ persentil 80 (top 20% paling laku)
- Margin ≥ 15%
- Sisa stok > 0

Scoring:
```
Score = 0.6 × normalize(Qty_Terjual) + 0.4 × normalize(Margin_pct)
```

## Konfigurasi

Edit `config.py`:

| Parameter | Default | Arti |
|---|---|---|
| `QTY_PERCENTILE` | `0.80` | Persentil qty untuk kandidat |
| `MARGIN_THRESHOLD_KANDIDAT` | `15.0` | Margin minimum (%) untuk kandidat |
| `MARGIN_BORDERLINE_MIN/MAX` | `0.0 / 5.0` | Range margin borderline (%) |
| `TARGET_MARGIN_KOREKSI` | `0.15` | Target margin rekomendasi harga koreksi |
| `PRICE_SCENARIOS` | `[0.10, 0.15, 0.20]` | Skenario kenaikan harga |
| `CHINA_KEYWORDS` | `["ocistok", "martkita", ...]` | Keyword untuk klasifikasi supplier China |
| `MARKET_KEYWORDS` | `["shopee ", ...]` | Keyword untuk klasifikasi supplier Market |
| `HPP_VARIANCE_THRESHOLD` | `0.15` | CV threshold (15%) untuk flag HPP volatile |
| `EXCLUDED_SKUS` | `{"ITBISA-BUBBLE-WRAP"}` | SKU yang di-exclude |

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
✓ Aggregasi per SKU
✓ Membangun tabel analisa
  Qty threshold (p80): 806 pcs
✓ Menulis laporan ke output/Analisa_Penjualan_ITBisa_2026.xlsx
```

## Workflow Bulanan

1. Update `Jual_2026.xlsx` dari export Shopee/Tiktok terbaru
2. Update `Stok_2026.xlsx` dengan pembelian baru
3. Run `python main.py` (default current year)
4. Buka Excel output, baca sheet `00_Summary` dulu untuk overview cepat

## Troubleshooting

**`Tidak ada file Stok ditemukan`** — pastikan ada minimal 1 file `Stok_*.xlsx` di folder `data/`.

**`Kolom hilang`** — struktur Excel berubah. Cek nama kolom raw di Excel vs `config.py`.

**`Tidak ada data jual untuk tahun X`** — tahun X tidak punya transaksi setelah filtering. Cek `Tanggal Pesan` di file jual.

**Pandas warning soal data validation** — bisa diabaikan, limitasi openpyxl saat read xlsx dengan dropdown/data validation.

## Catatan

- **Standalone**: bisa dirun tanpa Claude/AI
- **Idempotent**: input sama → output sama (kecuali timestamp di Summary)
- **Memory-friendly**: ditest dengan 15,000+ baris transaksi
- **Yearly snapshot**: setiap run filter ke 1 tahun berdasarkan `--year` flag
