# ITBisa Shop — Sales Analysis Bot

Tool analisa penjualan bulanan/tahunan untuk **IT Bisa Shop** dari data export marketplace
(Shopee, TikTok Shop, CoD, + legacy Tokopedia/Bukalapak). Output berupa workbook Excel
multi-sheet: barang terlaris, penyumbang profit, barang rugi, kandidat naik harga, analisa
supplier, reorder, dan rekap stok per gudang.

## Setup

```bash
pip install -r requirements.txt   # pandas, openpyxl
```

Taruh file export di folder `data/` (nama bebas, dicocokkan via glob):

- Stok  : `*BisaStok*.xlsx`   (sheet `BisaStok`)
- Jual  : `*BisaJual*.xlsx`   (sheet `BisaJualShopee`, `BisaJualTiktok`, `BisaJualTokopedia`,
  `BisaJualBukalapak`, `BisaJualCoD` — sheet yang tidak ada otomatis dilewati)

Bila ada beberapa file yang cocok, **file terbaru** (berdasarkan nama) dipakai sebagai
"workbook berjalan" untuk ledger stok; semua file dipakai untuk histori jual & HPP.

## Cara pakai

```bash
python main.py --sales 2026     # laporan penjualan satu tahun
python main.py --sales          # semua tahun yang ada di data jual (= --sales all)
python main.py --reorder        # analisa reorder standalone
python main.py --ab-test        # analisa A/B test perubahan harga (baca data/ab_tests.xlsx)
python main.py --all            # sales semua tahun + reorder + A/B test sekaligus
```

Hasil ditulis ke folder `output/`.

## Konsep penting

### HPP_WA vs HPP dasar harga (basis penetapan harga)

Ada **dua** angka HPP yang berbeda peran:

| | HPP_WA | HPP dasar harga (`hpp_pricing`) |
|---|---|---|
| Definisi | Weighted average semua pembelian relevan | Harga **lot luar negeri terakhir** (by Tanggal Bayar); fallback ke HPP_WA bila SKU tidak punya pembelian LN |
| Dipakai untuk | **Profit & Margin** (kebenaran P&L / realisasi) | **Analisa harga**: Markup %, Kandidat Naik Harga, Borderline, rekomendasi floor di Barang Rugi |

Alasan: harga supplier luar negeri (impor langsung) konsisten, jadi untuk pertanyaan
"apakah harus naik harga?" basis yang relevan adalah **biaya pengganti terkini** (harga lot
LN terakhir), bukan rata-rata tertimbang yang masih terbebani lot lama yang mahal.
Sebaliknya, perhitungan profit/margin tetap memakai HPP_WA karena itu merepresentasikan
biaya yang benar-benar terealisasi.

- **Luar negeri** didefinisikan ketat: hanya baris dengan kolom `Luar Negeri? = 1`
  (bukan tebakan dari nama supplier).
- Bila satu SKU punya beberapa lot LN, dipakai **lot terakhir** (Tanggal Bayar paling baru).
- Sumber HPP harga per SKU bisa dilihat di Sheet 07 kolom **"Sumber HPP Harga"**
  (`LN-terakhir` atau `WA`), bersanding dengan **"HPP Dasar Harga"** dan **"HPP/Buah (P&L)"**.

### Markup, bukan margin

Aturan harga minimum = **HPP dasar harga × 1.30** (markup 30% di atas HPP), **bukan**
gross margin 30%. Konstanta: `TARGET_MARKUP_KOREKSI = 0.30`, `MARKUP_THRESHOLD_KANDIDAT = 30.0`.

### Migrasi (carry-over stok)

Baris dengan kolom Toko berawalan `Migrasi` = stok akhir tahun yang dibawa ke file tahun
berikutnya — **referensi saja, bukan pembelian riil**. Otomatis di-drop saat ada data
pembelian non-Migrasi untuk SKU yang sama (mencegah double-count HPP & sisa stok berlebih).

## Struktur sheet output (laporan penjualan)

| Sheet | Isi |
|---|---|
| 00_Summary | Ringkasan total omzet, profit, margin |
| 01_Paling_Diminati | Barang dengan avg qty/order tertinggi |
| 02_Profit_Tertinggi | Penyumbang profit terbesar |
| 03_Barang_Rugi | Jual di bawah modal (rugi dari HPP_WA; rekomendasi dari HPP dasar harga) |
| 04_Margin_Borderline | Markup < 30% di bawah floor — wajib review |
| 05_Kandidat_Naik_Harga | Top qty + markup sehat → kandidat naik harga |
| 06_Per_Platform | Breakdown margin & biaya per marketplace |
| 07_Data_Lengkap_per_SKU | Semua metrik per SKU (termasuk HPP dasar harga & sumbernya) |
| 08_Supplier_Analysis | Analisa supplier (China vs lokal) |
| 09_Reorder_Analysis | Status reorder per SKU |
| 10_Reorder_Data_Lengkap | Detail perhitungan reorder |
| 11_Rekap_Stok_per_Gudang | Saldo stok per gudang |

## Struktur kode

- `config.py` — konstanta (glob, sheet, threshold, prefix Migrasi)
- `data_loader.py` — baca & bersihkan data stok/jual
- `analysis.py` — HPP_WA, HPP dasar harga (LN terakhir), agregasi per SKU, profit, reorder
- `tables.py` — bangun tabel-tabel analisa
- `excel_writer.py` — render workbook Excel
- `ab_testing.py` — analisa A/B test perubahan harga
- `main.py` — CLI entry point
