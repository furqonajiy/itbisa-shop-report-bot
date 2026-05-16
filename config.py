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
GREEN_FILL_COLOR = "C6EFCE"
YELLOW_FILL_COLOR = "FFEB9C"
LIGHT_GRAY_COLOR = "F2F2F2"
TITLE_COLOR = "1F4E78"
ALERT_TEXT_COLOR = "C00000"

FMT_RP = '"Rp"#,##0;[Red]("Rp"#,##0);"-"'
FMT_NUM = '#,##0;[Red](#,##0);"-"'
FMT_PCT = '0.0%;[Red](0.0%);"-"'
FMT_DEC = "#,##0.0"

# Source column names (raw from Excel with embedded newlines)
COL_STOK_SKU = "SKU"
COL_STOK_QTY = "Banyak\nBarang\n(Buah)"
COL_STOK_TOTAL_HPP = "Total\nHPP\n(Rp)"
COL_STOK_TANGGAL_BAYAR = "Tanggal\nBayar"
COL_STOK_TOKO = "Toko[spasi]Akun Pemesan"
COL_STOK_LUAR_NEGERI = "Luar\nNegeri?"

COL_JUAL_INVOICE = "Invoice"
COL_JUAL_VOID = "Void"
COL_JUAL_QTY = "Banyak\nTerjual\n(Buah)"
COL_JUAL_OMZET = "Omzet\nBarang\n(Rp)"
COL_JUAL_KODE_UNIK = "Kode\nUnik\n(Rp)"
COL_JUAL_TAMBAHAN = "Tambahan"
COL_JUAL_TANGGAL = "Tanggal\nPesan"
COL_JUAL_AKUN = "Akun\nPenjual"
