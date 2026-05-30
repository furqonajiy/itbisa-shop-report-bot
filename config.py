"""Configuration constants for ITBisa Sales Analysis."""
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent
DATA_DIR = PROJECT_ROOT / "data"
OUTPUT_DIR = PROJECT_ROOT / "output"

STOK_GLOB = "Stok_*.xlsx"
JUAL_GLOB = "Jual_*.xlsx"
OUTPUT_FILENAME = "Analisa_Penjualan_ITBisa_{year}.xlsx"

STOK_SHEET = "BisaStok"

# All supported jual sheets. Sheets not present in a file are skipped.
JUAL_SHEETS = [
    "BisaJualShopee",
    "BisaJualTiktok",
    "BisaJualTokopedia",
    "BisaJualBukalapak",
    "BisaJualCoD",
]
REQUIRED_JUAL_SHEET = "BisaJualShopee"

EXCLUDED_SKUS = {"ITBISA-BUBBLE-WRAP"}

# Inventory carry-over marker: toko names starting with this prefix represent
# year-end stock migrated to the new year file, NOT real purchases.
# Excluded from HPP_WA when matching real purchase data exists for the same SKU.
MIGRASI_PREFIX = "Migrasi"

QTY_PERCENTILE = 0.80
MARKUP_THRESHOLD_KANDIDAT = 30.0
MARKUP_BORDERLINE_MIN = 0.0
MARKUP_BORDERLINE_MAX = 30.0
TARGET_MARKUP_KOREKSI = 0.30

SCORE_WEIGHT_VELOCITY = 0.6
SCORE_WEIGHT_MARGIN = 0.4

PRICE_SCENARIOS = [0.10, 0.15, 0.20]

TOP_N_DIMINATI = 40
TOP_N_PROFIT = 40
TOP_N_PER_PLATFORM = 10

# Supplier classification keywords (case-insensitive substring match on supplier name)
CHINA_KEYWORDS = ["ocistok", "martkita", "aliexpress", "jasa impor", "1688", "alibaba"]
MARKET_KEYWORDS = ["shopee ", "tokopedia ", "bukalapak ", "tiktok "]
HPP_VARIANCE_THRESHOLD = 0.15
SUPPLIER_TOP_N_SINGLE_SOURCE = 15

FONT_NAME = "Arial"
HEADER_BG_COLOR = "1F4E78"
HEADER_TEXT_COLOR = "FFFFFF"
RED_FILL_COLOR = "F8CBAD"
ORANGE_FILL_COLOR = "FCD5B4"
GREEN_FILL_COLOR = "C6EFCE"
YELLOW_FILL_COLOR = "FFEB9C"
BLUE_FILL_COLOR = "DDEBF7"
LIGHT_GRAY_COLOR = "F2F2F2"
TITLE_COLOR = "1F4E78"
ALERT_TEXT_COLOR = "C00000"

FMT_RP = '"Rp"#,##0;[Red]("Rp"#,##0);"-"'
FMT_NUM = '#,##0;[Red](#,##0);"-"'
FMT_PCT = '0.0%;[Red](0.0%);"-"'
FMT_DEC = "#,##0.0"

# A/B Testing config
AB_TESTS_FILENAME = "ab_tests.xlsx"
AB_TESTS_SHEET = "BisaABTest"
AB_TESTS_OUTPUT_FILENAME = "Analisa_AB_Test.xlsx"
COL_AB_SKU = "SKU"
COL_AB_TANGGAL = "Tanggal Perubahan"
COL_AB_NAMA = "Nama Test"
COL_AB_CATATAN = "Catatan"
AB_MIN_DAYS_POST = 3

# Reorder analysis config
REORDER_OUTPUT_FILENAME = "Analisa_Reorder.xlsx"
LEAD_TIME_CHINA_MONTHS = 2.0
LEAD_TIME_MARKET_MONTHS = 0.25
BULK_CHINA_SHARE_THRESHOLD = 0.5
VELOCITY_MIN_ACTIVE_MONTHS = 3
CV_STABLE_MAX = 0.3
CV_MODERATE_MAX = 0.7
SAFETY_MULT_STABLE = 1.2
SAFETY_MULT_MODERATE = 1.5
SAFETY_MULT_VOLATILE = 2.0
TARGET_MONTHS_POST_REORDER = 6
SLOW_DEAD_MAX_VELOCITY = 0.5
ROP_URGENT_RATIO = 0.7
ROP_NOW_RATIO = 1.0
ROP_SOON_RATIO = 1.3
OVERSTOCK_MONTHS = 12.0

# Source column names (raw from Excel with embedded newlines)
COL_STOK_SKU = "SKU"
COL_STOK_QTY = "Banyak\nBarang\n(Buah)"
COL_STOK_TOTAL_HPP = "Total\nHPP\n(Rp)"
COL_STOK_TANGGAL_BAYAR = "Tanggal\nBayar"
COL_STOK_TOKO = "Toko[spasi]Akun Pemesan"
COL_STOK_LUAR_NEGERI = "Luar\nNegeri?"

# --- Current-workbook stock ledger (reconcile to BisaRekapBarang) ---
# sisa_stok is computed from the latest stok+jual workbook (by filename), using the
# same formula as the Google Sheets rekap: arrived beli − nonvoid jual + ketemu
# − hilang ± pindah, per (SKU, gudang). Migrasi rows are KEPT (they are the opening
# balance). No dedup (the rekap sums every row).
COL_STOK_TANGGAL_SAMPAI = "Tanggal\nSampai"   # filled = "sudah sampai" (arrived)
COL_STOK_ALAMAT = "Alamat Pengiriman"          # destination gudang for a purchase
COL_JUAL_GUDANG = "Lokasi Gudang"              # gudang a sale is deducted from

HILANG_SHEET = "BisaHilang"
COL_HILANG_SKU = "SKU"
COL_HILANG_KETEMU = "Banyak\nKetemu"           # found (+)
COL_HILANG_HILANG = "Banyak\nHilang"           # lost (−)
COL_HILANG_GUDANG = "Lokasi Gudang"

PINDAH_SHEET = "BisaPindahBarang"
COL_PINDAH_SKU = "Unnamed: 1"                  # SKU is col B; header is blank
COL_PINDAH_TAMBAH = "Lokasi\nPenambahan\nBarang"   # gudang receiving (+)
COL_PINDAH_KURANG = "Lokasi\nPengurangan\nBarang"  # gudang losing (−)
COL_PINDAH_QTY = "Banyak Barang"

# Ledger jual scope: all sheets starting with this prefix in the current jual file
# (matches BisaRekapBarang, which includes Blibli/Investasi beyond JUAL_SHEETS).
LEDGER_JUAL_PREFIX = "BisaJual"
LEDGER_SHEET_NAME = "11_Rekap_Stok_per_Gudang"

COL_JUAL_INVOICE = "Invoice"
COL_JUAL_VOID = "Void"
COL_JUAL_QTY = "Banyak\nTerjual\n(Buah)"
COL_JUAL_OMZET = "Omzet\nBarang\n(Rp)"
COL_JUAL_KODE_UNIK = "Kode\nUnik\n(Rp)"
COL_JUAL_TAMBAHAN = "Tambahan"
COL_JUAL_TANGGAL = "Tanggal\nPesan"
COL_JUAL_AKUN = "Akun\nPenjual"
