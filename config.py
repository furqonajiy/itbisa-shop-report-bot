"""Configuration constants for ITBisa Sales Analysis."""
from pathlib import Path

# Paths
PROJECT_ROOT = Path(__file__).parent
DATA_DIR = PROJECT_ROOT / "data"
OUTPUT_DIR = PROJECT_ROOT / "output"

# File templates
STOK_FILENAME = "Stok_{year}.xlsx"
JUAL_FILENAME = "Jual_{year}.xlsx"
OUTPUT_FILENAME = "Analisa_Penjualan_ITBisa_{year}.xlsx"

# Sheet names
STOK_SHEET = "BisaStok"
JUAL_SHOPEE_SHEET = "BisaJualShopee"
JUAL_TIKTOK_SHEET = "BisaJualTiktok"

# Excluded SKUs (packaging, etc.)
EXCLUDED_SKUS = {"ITBISA-BUBBLE-WRAP"}

# Analysis thresholds
QTY_PERCENTILE = 0.80
MARGIN_THRESHOLD_KANDIDAT = 15.0
MARGIN_BORDERLINE_MIN = 0.0
MARGIN_BORDERLINE_MAX = 5.0
TARGET_MARGIN_KOREKSI = 0.15

# Score weights for "kandidat naik harga"
SCORE_WEIGHT_VELOCITY = 0.6
SCORE_WEIGHT_MARGIN = 0.4

# Scenario prices (percentage increase)
PRICE_SCENARIOS = [0.10, 0.15, 0.20]

# Top N for tables
TOP_N_DIMINATI = 40
TOP_N_PROFIT = 40
TOP_N_PER_PLATFORM = 10

# Excel styling
FONT_NAME = "Arial"
HEADER_BG_COLOR = "1F4E78"
HEADER_TEXT_COLOR = "FFFFFF"
RED_FILL_COLOR = "F8CBAD"
GREEN_FILL_COLOR = "C6EFCE"
YELLOW_FILL_COLOR = "FFEB9C"
LIGHT_GRAY_COLOR = "F2F2F2"
TITLE_COLOR = "1F4E78"
ALERT_TEXT_COLOR = "C00000"

# Number formats
FMT_RP = '"Rp"#,##0;[Red]("Rp"#,##0);"-"'
FMT_NUM = '#,##0;[Red](#,##0);"-"'
FMT_PCT = '0.0%;[Red](0.0%);"-"'
FMT_DEC = "#,##0.0"

# Source column names (raw from Excel files, with embedded newlines)
COL_STOK_SKU = "SKU"
COL_STOK_QTY = "Banyak\nBarang\n(Buah)"
COL_STOK_TOTAL_HPP = "Total\nHPP\n(Rp)"
COL_STOK_TANGGAL_BAYAR = "Tanggal\nBayar"

COL_JUAL_SKU = "SKU"
COL_JUAL_INVOICE = "Invoice"
COL_JUAL_QTY = "Banyak\nTerjual\n(Buah)"
COL_JUAL_OMZET = "Omzet\nBarang\n(Rp)"
COL_JUAL_VOID = "Void"
COL_JUAL_KODE_UNIK = "Kode\nUnik\n(Rp)"
COL_JUAL_TAMBAHAN = "Tambahan"
COL_JUAL_TANGGAL = "Tanggal\nPesan"
COL_JUAL_AKUN = "Akun\nPenjual"
