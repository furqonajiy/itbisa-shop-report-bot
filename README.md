# ITBisa Shop Report Bot

Tool standalone untuk generate laporan analisa penjualan ITBisa Shop dari data Shopee, Tokopedia, dan TikTok Shop. Dirancang untuk dirun berulang tanpa memerlukan Claude/internet.

## Setup

Persyaratan: **Python 3.10+**

```bash
cd itbisa-shop-report-bot
pip install -r requirements.txt
```

## Cara Pakai

### 1. Siapkan data

Tempatkan file Excel di folder `data/`:

- `Stok_<tahun>.xlsx` — file stock dengan sheet `BisaStok` (headers di baris ke-2, kolom kategori di baris ke-1)
- `Jual_<tahun>.xlsx` — file penjualan dengan sheet `BisaJualShopee` (wajib) dan `BisaJualTiktok` (opsional, berisi Tokopedia + Tiktok)

Contoh struktur:
```
data/
├── Stok_2026.xlsx
└── Jual_2026.xlsx
```

### 2. Jalankan analisa

```bash
# Analisa tahun berjalan (default: tahun sekarang)
python main.py

# Analisa tahun spesifik
python main.py --year 2024

# Custom path
python main.py --year 2024 --data-dir /path/to/data --output-dir /path/to/output
```

### 3. Lihat hasil

File output ada di folder `output/`:
- `Analisa_Penjualan_ITBisa_<tahun>.xlsx` — 8 sheet lengkap

## Struktur Output Excel

| Sheet | Isi |
|-------|-----|
| `00_Summary` | Ringkasan performa + temuan utama (data-driven narrative) |
| `01_Paling_Diminati` | Top 40 SKU by Qty Terjual |
| `02_Profit_Tertinggi` | Top 40 SKU by Total Profit |
| `03_Barang_Rugi` | SKU rugi + rekomendasi harga koreksi |
| `04_Margin_Borderline` | SKU margin 0-5% (rawan rugi) |
| `05_Kandidat_Naik_Harga` | SKU rekomendasi naik harga + skenario +10%/+15%/+20% |
| `06_Per_Platform` | Breakdown per Shopee/Tokopedia/Tiktok + top SKU per platform |
| `07_Data_Lengkap_per_SKU` | Full per-SKU table untuk drill-down manual |

## Struktur Project

```
itbisa-shop-report-bot/
├── README.md             # this file
├── requirements.txt      # Python dependencies
├── .gitignore
├── config.py             # constants (paths, thresholds, colors, column names)
├── data_loader.py        # load & clean stok/jual Excel files
├── analysis.py           # HPP weighted average + profit + per-SKU aggregation
├── tables.py             # table builders (diminati, profit, rugi, kandidat, dll.)
├── excel_writer.py       # Excel output with styling helpers
├── main.py               # entry point with CLI
├── data/                 # input files (gitignored)
└── output/               # generated reports
```

## Metodologi Analisa

### HPP (Harga Pokok Penjualan)
**Weighted Average** dari semua pembelian historis per SKU:
```
HPP_WA = sum(Total_HPP_Rp) / sum(Banyak_Barang_Buah)  per SKU
```

### Profit per Transaksi
```
Untung = Omzet − (HPP_WA × Qty_Jual) + Tambahan + Kode_Unik
```

`Tambahan` dan `Kode_Unik` sudah signed di source data (negatif = biaya admin). Tidak perlu di-negate lagi.

### Exclusions Otomatis (data jual)
1. Baris dengan SKU null
2. Invoice diawali `"Dummy"` (test data)
3. `Void = True` (pesanan dibatalkan)
4. SKU di `EXCLUDED_SKUS` set (default: `ITBISA-BUBBLE-WRAP`)
5. Baris dengan numeric invalid (header row yang ke-scatter di tengah data)

SKU yang dijual tapi tidak ada HPP-nya di stok: di-warning ke console dan di-exclude dari analisa profit.

### Kriteria "Kandidat Naik Harga"
SKU yang lolos **3 filter**:
- Qty terjual ≥ persentil 80 (top 20% paling laku)
- Margin saat ini ≥ 15%
- Sisa stok > 0

Lalu di-scoring dengan formula:
```
Score = 0.6 × normalize(Qty_Terjual) + 0.4 × normalize(Margin_pct)
```

## Konfigurasi Threshold

Edit `config.py` untuk mengubah parameter analisa tanpa mengubah logic:

| Parameter | Default | Arti |
|-----------|---------|------|
| `QTY_PERCENTILE` | `0.80` | Persentil qty terjual untuk kandidat naik harga |
| `MARGIN_THRESHOLD_KANDIDAT` | `15.0` | Margin minimum (%) untuk kandidat |
| `MARGIN_BORDERLINE_MIN/MAX` | `0.0 / 5.0` | Range margin borderline (%) |
| `TARGET_MARGIN_KOREKSI` | `0.15` | Target margin untuk rekomendasi harga koreksi |
| `SCORE_WEIGHT_VELOCITY/MARGIN` | `0.6 / 0.4` | Bobot scoring |
| `PRICE_SCENARIOS` | `[0.10, 0.15, 0.20]` | Skenario kenaikan harga |
| `TOP_N_DIMINATI` / `TOP_N_PROFIT` | `40 / 40` | Top N untuk tabel |
| `EXCLUDED_SKUS` | `{"ITBISA-BUBBLE-WRAP"}` | SKU yang di-exclude |

## Console Logging

Script menulis log Bahasa Indonesia ke console:

```
✓ Membaca file stok: Stok_2026.xlsx
  Stok valid rows: 186 (dari 261 mentah)
✓ Membaca file jual: Jual_2026.xlsx
  Loaded BisaJualShopee: 2025 rows
  Loaded BisaJualTiktok: 682 rows
✓ Cleaning jual:
  - Total mentah         :   2707
  - Dibuang (Dummy)      :      4
  - Dibuang (invalid)    :     38
  - Total bersih         :   2663
✓ HPP weighted average dihitung untuk 132 SKU
⚠ 2 SKU dijual tanpa data HPP (di-exclude, 6 pcs):
    - ITBISA-NPN-2N2222-TO92
    - 20aPCS-ITBISA-SOCKET-IC-DIP28-NARROW
✓ Menulis laporan ke output/Analisa_Penjualan_ITBisa_2026.xlsx
```

## Troubleshooting

**`FileNotFoundError`**: pastikan file Excel ada di `data/` dengan nama yang persis sesuai (`Stok_2026.xlsx`, `Jual_2026.xlsx`).

**`Kolom hilang di sheet...`**: struktur Excel berubah. Cek nama kolom di file vs `config.py` (kolom raw dengan `\n` newlines).

**Pandas warning soal data validation**: bisa diabaikan, ini limitasi openpyxl untuk fitur data validation di Excel yang tidak relevan untuk read-only access.

## Catatan Pengembangan

- File ini **standalone** — bisa dirun tanpa Claude/AI
- **Idempotent** — dirun berulang menghasilkan output sama (selama input sama)
- **Memory-friendly** — sudah ditest dengan ribuan baris transaksi
- Untuk monthly use: file `Jual_<tahun>.xlsx` di-update terus-menerus, script analisanya membaca state terbaru saat dirun
