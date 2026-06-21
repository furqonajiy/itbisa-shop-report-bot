# Cara Pakai — ITBisa Shop Report Bot

Panduan praktis untuk menjalankan tool ini di **Windows (PowerShell)**. Untuk detail teknis
tiap laporan, lihat `README.md`.

Tool ini jalan **offline** di komputer sendiri: tidak ada koneksi internet, tidak ada API,
tidak ada token. Kamu kasih file Excel export, dia balikin workbook analisa.

---

## 1. Persiapan (sekali saja)

1. Pastikan **Python 3.10+** sudah terpasang. Cek di PowerShell:
   ```powershell
   python --version
   ```
2. Dari folder project, pasang dependency-nya:
   ```powershell
   pip install -r requirements.txt
   ```
   (yang dipasang cuma `pandas` dan `openpyxl`.)

---

## 2. Siapkan data

Taruh file export di folder **`data/`**. Nama file bebas, asal **mengandung kata kunci** ini:

| Jenis | Pola nama file | Sheet yang dibutuhkan |
|---|---|---|
| Pembelian / stok | `*BisaStok*.xlsx` | `BisaStok` (file terbaru juga butuh `BisaHilang` + `BisaPindahBarang`) |
| Penjualan | `*BisaJual*.xlsx` | minimal `BisaJualShopee` (sheet `BisaJual*` lain ikut kalau ada) |

Catatan:
- Boleh banyak file (mis. per tahun). Semua dipakai untuk histori penjualan & HPP.
- File **terbaru** (urut nama) dipakai sebagai "current workbook" untuk hitung **sisa stok**.
- Kolom **`Toko`** di `BisaStok` = supplier/forwarder yang sudah distandardisasi
  (mis. `Ocistok/Martkita`, `AliExpress`, `Shopee`); kolom **`Luar Negeri?`** = penanda
  barang impor. Pastikan dua kolom ini terisi rapi supaya analisa reorder & supplier akurat.

---

## 3. Jalankan

Cara paling gampang — jalankan **semuanya** sekaligus:

```powershell
python main.py
```

Hasilnya muncul di folder **`output/`**. Selesai.

Kalau cuma butuh satu laporan, pakai flag-nya:

```powershell
python main.py --sales 2026     # laporan penjualan satu tahun
python main.py --sales          # laporan penjualan semua tahun
python main.py --trend          # tren & musiman lintas tahun
python main.py --reorder        # rekomendasi restock + rekap stok per gudang
python main.py --cashflow       # rencana budget belanja restock per bulan & supplier
python main.py --deadstock      # modal beku di stok lambat/mati + cara membebaskannya
python main.py --ab-test        # analisa hasil ubah harga (A/B test)
python main.py --restock-check  # cek harga restock vs harga jual per marketplace
```

Mau output di folder lain? Pakai `--data-dir` / `--output-dir`:

```powershell
python main.py --data-dir "D:\data" --output-dir "D:\hasil"
```

---

## 4. Apa saja yang dihasilkan (7 laporan)

| File di `output/` | Untuk apa | Sheet penting |
|---|---|---|
| `Analisa_Penjualan_ITBisa_<tahun>.xlsx` | Histori penjualan per tahun: apa yang laku, untung, rugi, kandidat naik harga | `01_Paling_Diminati`, `02_Profit_Tertinggi`, `03_Barang_Rugi`, `05_Kandidat_Naik_Harga` |
| `Analisa_Tren_Musiman.xlsx` | Lagi tumbuh atau tidak? Bulan apa yang selalu ramai? | `01_Tren_Tahunan`, `03_Musiman` |
| `Analisa_Reorder.xlsx` | Barang apa yang harus dibeli sekarang & berapa banyak | `01_Reorder_Action`, `03_Rekap_Stok_per_Gudang` |
| `Analisa_Cashflow_Restock.xlsx` | Butuh modal restock berapa, kapan, ke supplier mana | `01_Kalender_per_Bulan`, `02_Detail_per_SKU` |
| `Analisa_Modal_Beku.xlsx` | Modal yang nyangkut di stok lambat/mati + cara cairin | `01_Modal_Beku_per_SKU`, `02_Per_Supplier` |
| `Analisa_AB_Test.xlsx` | Apakah perubahan harga benar-benar berhasil | (opsional — lihat bagian 5) |
| `Analisa_Restock_Check.xlsx` | Layak tidak restock di harga China ini & jual berapa | (opsional — lihat bagian 5) |

---

## 5. Dua laporan opsional (perlu diisi dulu)

`--ab-test` dan `--restock-check` butuh template yang kamu isi sendiri. Saat menjalankan
`python main.py` (semua), kedua langkah ini **dilewati otomatis** kalau template-nya kosong —
jadi proses tidak pernah macet.

Cara mengaktifkan:

1. Bikin template-nya (cukup sekali):
   ```powershell
   python main.py --ab-test          # bikin data/ab_tests.xlsx kalau belum ada
   python main.py --restock-check    # bikin data/restock_check.xlsx kalau belum ada
   ```
2. Buka file template di `data/`, isi barisnya:
   - **`ab_tests.xlsx`** — catat tiap kali kamu ubah harga sebuah SKU (SKU + tanggal + harga
     lama/baru). Setelah ada data penjualan di harga baru, laporannya menilai dampaknya.
   - **`restock_check.xlsx`** — isi SKU + `Toko` + `Harga RMB` dan/atau `HPP IDR`, plus rentang
     harga kompetitor (`Kompetitor Min`/`Max`). Tool memprediksi HPP mendarat dan rekomendasi
     harga jual per marketplace.
3. Jalankan lagi (`python main.py` atau flag-nya), laporan akan ikut terbentuk.

---

## 6. Rutinitas yang disarankan

- **Tiap bulan:** export data baru → `python main.py` → buka **Reorder** (beli yang
  STOCKOUT/URGENT) → cek **Cashflow** untuk tagihannya.
- **Jelang musim ramai (Agustus–September):** lihat **Tren/Musiman** untuk stok jelang
  puncak Oktober–Januari.
- **Tiap kuartal:** buka **Modal Beku** (cairin stok yang nyangkut), dan **Penjualan
  `05_Kandidat_Naik_Harga`** → naikkan beberapa harga → catat di `ab_tests.xlsx` → cek
  hasilnya lewat **A/B Test**.

---

## 7. Masalah umum

- **"No files found" / laporan kosong** → cek nama file di `data/` sudah mengandung
  `BisaStok` / `BisaJual` dan ekstensinya `.xlsx`.
- **Sisa stok terlihat aneh / minus (OVERSOLD)** → tool menandai SKU OVERSOLD di console dan
  di `00_Summary`; biasanya karena ada penjualan tanpa data pembelian/migrasi yang cocok.
- **A/B test / restock-check tidak muncul** → templatenya masih kosong; isi dulu (lihat
  bagian 5).
- **`python` tidak dikenali** → Python belum terpasang atau belum masuk PATH; pasang ulang
  dan centang "Add Python to PATH".
