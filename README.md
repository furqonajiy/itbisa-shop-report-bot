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

### Harga Sekarang (baseline analisa naik harga)

Kolom **"Harga Sekarang"** (Sheet 04 & 05, juga Sheet 07) = **harga satuan terendah
(Omzet/Qty) pada hari penjualan non-CoD terakhir** suatu SKU.

- Bukan rata-rata seumur hidup (`harga_jual_avg`) yang tercampur harga lama; ambil dari
  **hari jual terakhir** supaya mewakili harga yang berlaku sekarang.
- **Minimum di hari itu** dipakai agar menangkap tier grosir terendah yang masih aktif,
  sekaligus mengabaikan (a) harga promo basi dari awal bulan dan (b) "undian" eceran vs
  grosir dari transaksi tunggal yang kebetulan paling akhir.
- **CoD dikecualikan** (harganya beda channel).
- **Markup %** (yang menentukan SKU masuk Kandidat ≥30% / Borderline <30%) dihitung dari
  `(Harga Sekarang − HPP dasar harga) / HPP dasar harga`, begitu juga skenario
  Harga +10/15/20% di sheet Kandidat.
- SKU tanpa penjualan non-CoD sama sekali → fallback ke `harga_jual_avg`.

### Guard "harga baru naik" (Kandidat Naik Harga)

`Harga Sekarang` itu titik-waktu (hari jual terakhir), tapi `Qty Terjual`/`Profit`
itu kumulatif. Kalau harga **baru saja dinaikkan**, qty & profit itu diraih di harga
**lama yang lebih murah** — jadi merekomendasikan naik lagi dan memproyeksikan profit
dari qty setahun penuh di harga baru itu **tidak valid**. Sheet 05 mendeteksi ini dan
**menahan** rekomendasinya (baris di-grey, kolom `Harga +10/15/20%` & `Proyeksi Profit`
dikosongkan, Saran diganti `⏳ Harga baru naik <tgl> (Rp lama→Rp baru); baru X% qty di
harga baru — kumpulkan data dulu, jangan naik lagi`).

- **Tanggal perubahan**: dari `ab_tests.xlsx` bila SKU tercatat (otoritatif), else
  **auto-deteksi** lonjakan harga (rata-rata tertimbang window terbaru vs baseline sebelumnya).
- **Ditandai hanya bila** (a) `Harga Sekarang` benar-benar di atas harga lama
  (≥ `PRICE_CHANGE_MIN_STEP`) **dan** (b) demand di harga baru masih tipis
  (< `PRICE_CHANGE_VALIDATION_MIN_SHARE` dari qty tahun itu). Harga yang sudah lama
  stabil (mis. Rp999 berbulan-bulan) **tidak** ditandai walau jauh di atas rata-rata.
- Tunable di `config.py`: `PRICE_CHANGE_RECENT_DAYS`, `PRICE_CHANGE_MIN_STEP`,
  `PRICE_CHANGE_VALIDATION_MIN_SHARE`, `PRICE_CHANGE_AUTO_RECENT_DAYS`,
  `PRICE_CHANGE_AUTO_PRIOR_DAYS`, `PRICE_CHANGE_PRE_WINDOW_DAYS`. Lihat
  `compute_price_change_status` di `analysis.py`.

### Markup, bukan margin

Aturan harga minimum = **HPP dasar harga × 1.30** (markup 30% di atas HPP), **bukan**
gross margin 30%. Konstanta: `TARGET_MARKUP_KOREKSI = 0.30`, `MARKUP_THRESHOLD_KANDIDAT = 30.0`.

### Migrasi (carry-over stok)

Baris dengan kolom Toko berawalan `Migrasi` = stok akhir tahun yang dibawa ke file tahun
berikutnya — **referensi saja, bukan pembelian riil**. Otomatis di-drop saat ada data
pembelian non-Migrasi untuk SKU yang sama (mencegah double-count HPP & sisa stok berlebih).

### Lead time reorder (per shop, dari data)

Titik reorder (ROP) butuh tahu **berapa lama barang sampai** supaya tidak kehabisan saat
menunggu kiriman. Lead time itu **sifat SHOP/forwarder, bukan per-SKU** — AliExpress
(~1 bln) jauh lebih cepat dari forwarder laut Ocistok/Martkita (~2,5 bln). Caranya:

1. **Per shop**: tiap shop impor dihitung **persentil 75** (`LEAD_TIME_PERCENTILE`) dari
   selisih `Tanggal Bayar` → `Tanggal Sampai` (baris Migrasi dikecualikan); shop dengan
   < `LEAD_TIME_MIN_LOTS` lot pakai persentil global impor. **Ocistok = Martkita = 1688
   dianggap satu forwarder** (Ocistok rebrand jadi Martkita; 1688 lewat mereka) —
   `OCISTOK_KEYWORDS` / `IMPORT_SHOP_KEYWORDS`.
2. **Per SKU**: ambil lead **shop paling lambat yang menyuplai ≥ `LEAD_SHOP_MIN_SHARE`
   qty** SKU itu (rencanakan untuk impor lambat, bukan top-up lokal sesekali). Forwarder
   diambil dari kolom **`Toko` yang sudah distandarkan** (mis. `Jasa Impor Tiongkok`,
   `Ocistok/Martkita`, `AliExpress`, `Osell`), sedangkan status impor dari share qty
   **`Luar Negeri?`/keyword China** (otoritatif); SKU impor di-floor ke lead global impor.
   SKU yang disuplai lokal pakai `LEAD_TIME_MARKET_MONTHS` (≈ 1 minggu).

> **Catatan data**: `Toko` = sumber/forwarder yang sudah distandarkan; detail akun
> bayar + invoice/resi pindah ke `Keterangan Pembelian`; `Luar Negeri?` tetap penanda
> impor (mis. "Tokopedia Furqonajiy" yang bayar Jasa Impor Tiongkok kini `Toko = Jasa
> Impor Tiongkok`). Loader otomatis menerima header lama `Toko[spasi]Akun Pemesan`.

Lihat `compute_lead_time_months` di `analysis.py`.

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

## A/B Test perubahan harga (`--ab-test`)

Mengukur **apakah kenaikan harga benar berpengaruh ke profit** — bukan sekadar pre vs post,
karena profit bergerak oleh banyak hal (tren, musiman, stok habis). Yang dibandingkan:

- **Window pre setara**: `AB_PRE_WINDOW_DAYS` hari (default 60) sebelum tanggal perubahan,
  bukan rata-rata seumur hidup (baseline all-time bikin delta menggelembung). Semua metrik
  pakai daily-rate (per hari) supaya adil meski panjang window beda.
- **Profit bridge**: `Δprofit/hari = Efek Harga + Efek Volume + Interaksi + Efek Admin`.
  Memisahkan tambahan margin dari kenaikan harga (efek harga, +) terhadap dampak perubahan
  volume (efek volume). Kalau efek volume negatif besar dan mengalahkan efek harga → kenaikan
  harga merugikan.
- **Break-even turun qty** = `1 − (margin_pre / margin_post)`: berapa % volume boleh hilang
  sebelum profit balik ke level sebelumnya. **Headroom** = jarak qty aktual ke batas itu
  (+ = aman).
- **Elastisitas** = %Δqty ÷ %Δprice (diagnostik). Bila **positif** (qty & harga sama-sama
  naik) → ada faktor lain; efek harga tak bisa diisolasi → ditandai di kolom Catatan.
- **Flag confound**: post terlalu pendek/sedikit transaksi, baseline pre tipis, qty post
  didominasi 1 order grosir, elastisitas positif.
- **Verdict** deskriptif (✅/🟡/🔴/⚪) berbasis arah profit + break-even, diturunkan ke 🟡 bila
  atribusi lemah.

Belum termasuk Difference-in-Differences (kontrol SKU lain) & bootstrap CI — langkah lanjutan.

Config `data/ab_tests.xlsx` (sheet `BisaABTest`): `SKU`, `Tanggal Perubahan`, `Nama Test`,
`Catatan`. Output `output/Analisa_AB_Test.xlsx` (sheet `00_Summary`, `01_Test_Results`).

## Struktur kode

- `config.py` — konstanta (glob, sheet, threshold, prefix Migrasi)
- `data_loader.py` — baca & bersihkan data stok/jual
- `analysis.py` — HPP_WA, HPP dasar harga (LN terakhir), agregasi per SKU, profit, reorder
- `tables.py` — bangun tabel-tabel analisa
- `excel_writer.py` — render workbook Excel
- `ab_testing.py` — analisa A/B test perubahan harga
- `main.py` — CLI entry point
