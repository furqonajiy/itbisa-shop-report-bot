"""Configuration constants for ITBisa Sales Analysis."""
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent
DATA_DIR = PROJECT_ROOT / "data"
OUTPUT_DIR = PROJECT_ROOT / "output"

STOK_GLOB = "*BisaStok*.xlsx"
JUAL_GLOB = "*BisaJual*.xlsx"
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

# --- Recent price-change guard for Kandidat Naik Harga (sheet 05) ---
# A SKU whose 'harga sekarang' is a RECENT price increase that little demand has
# been observed at yet is NOT a valid "raise further" candidate: its qty/profit
# were earned at the old, lower price, so the +% suggestions and profit projection
# would extrapolate stale demand. Detected via ab_tests.xlsx change dates
# (authoritative) + automatic two-window price-step detection (fallback). Flagged
# rows stay listed but their price/projection columns are blanked and the Saran is
# replaced with a "kumpulkan data dulu" note.
PRICE_CHANGE_RECENT_DAYS = 75          # only changes this recent count as "baru naik"
PRICE_CHANGE_MIN_STEP = 0.05           # harga_sekarang must be ≥5% over old price
PRICE_CHANGE_VALIDATION_MIN_SHARE = 0.25  # flag if <25% of year qty sold post-change
PRICE_CHANGE_AUTO_RECENT_DAYS = 30     # auto-detect: recent window (qty-weighted avg)
PRICE_CHANGE_AUTO_PRIOR_DAYS = 90      # auto-detect: prior baseline window
PRICE_CHANGE_PRE_WINDOW_DAYS = 60      # window before change date → harga_lama

SCORE_WEIGHT_VELOCITY = 0.6
SCORE_WEIGHT_MARGIN = 0.4

PRICE_SCENARIOS = [0.10, 0.15, 0.20]

TOP_N_DIMINATI = 40
TOP_N_PROFIT = 40
TOP_N_PER_PLATFORM = 10

# Supplier classification keywords (case-insensitive substring match on the
# standardized Toko value, e.g. "Ocistok/Martkita", "AliExpress", "Shopee").
CHINA_KEYWORDS = ["ocistok", "martkita", "aliexpress", "jasa impor", "1688", "alibaba", "osell"]
# Consistent China-direct forwarder. Ocistok rebranded to Martkita (SAME company),
# and 1688 orders are fulfilled through them — so all three are ONE shop/channel.
OCISTOK_KEYWORDS = ["ocistok", "martkita", "1688"]
# Lead time is a per-SHOP (forwarder) property, not per-SKU — AliExpress ships
# faster than the Ocistok/Martkita sea-freight forwarder. Each import shop gets its
# own observed lead (Tanggal Bayar→Sampai); a SKU inherits the lead of the shop that
# supplies most of its purchase qty. Anything not listed here (local marketplaces,
# domestic distributors) is treated as "Local" and ships in days (LEAD_TIME_MARKET_MONTHS).
IMPORT_SHOP_KEYWORDS = {
    "Ocistok/Martkita": OCISTOK_KEYWORDS,   # one forwarder (Ocistok=Martkita; 1688 via them)
    "AliExpress": ["aliexpress"],
    "Alibaba": ["alibaba"],
    "Jasa Impor": ["jasa impor"],
    "Osell": ["osell"],
}
# Match the standardized Toko marketplaces exactly (no trailing space — values are
# now clean "Shopee"/"Tokopedia", not "Shopee Aji").
MARKET_KEYWORDS = ["shopee", "tokopedia", "bukalapak", "blibli", "tiktok"]
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
# Window pre yang setara (hari sebelum tanggal perubahan). Baseline all-time
# membandingkan post pendek vs rata-rata bertahun-tahun → delta menggelembung.
AB_PRE_WINDOW_DAYS = 60
# Satu order > sekian fraksi qty post = qty didominasi 1 borongan → flag confound.
AB_BULK_CONCENTRATION = 0.40
# Min transaksi di window pre sebelum baseline dianggap cukup.
AB_MIN_TRANS_PRE = 3
# Min transaksi di window post sebelum klaim sebab-akibat dianggap cukup.
AB_MIN_TRANS_POST = 3

# Reorder analysis config
REORDER_OUTPUT_FILENAME = "Analisa_Reorder.xlsx"
LEAD_TIME_CHINA_MONTHS = 2.0
LEAD_TIME_MARKET_MONTHS = 0.25
# Observed lead time: derive each SKU's China shipping time from BisaStok
# (Tanggal Sampai − Tanggal Bayar) at this percentile so the reorder point
# survives typical delays (not just the median). Per-SKU when it has ≥ MIN_LOTS
# dated China lots, else the global-China percentile, else the constants above.
LEAD_TIME_PERCENTILE = 0.75
LEAD_TIME_MIN_LOTS = 2
LEAD_TIME_MAX_DAYS = 365          # ignore lots with impossible/garbage date gaps
# Mixed-sourcing: a SKU's lead = the SLOWEST shop supplying ≥ this qty share of it
# (so an item bought partly via the slow Ocistok forwarder is planned for that delay,
# not its occasional fast local top-up). Lower = more conservative / fewer stockouts.
LEAD_SHOP_MIN_SHARE = 0.20
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

# --- Restock price check (--restock-check) ---
# Evaluate a supplier's offered restock price for a SKU and recommend the selling
# price per marketplace. Input data/restock_check.xlsx (auto-template if missing).
RESTOCK_CHECK_FILENAME = "restock_check.xlsx"
RESTOCK_CHECK_SHEET = "RestockCheck"
RESTOCK_OUTPUT_FILENAME = "Analisa_Restock_Check.xlsx"
RESTOCK_PLATFORMS = ["Shopee", "Tokopedia", "Tiktok"]
# Target net profit = this fraction of landed HPP, realized AFTER the platform fee.
# So sell_min = HPP * (1 + markup) / (1 - fee). e.g. HPP 450, 30% → net ≥ 135.
RESTOCK_TARGET_NET_MARKUP = 0.30
RESTOCK_COST_TOL = 0.10               # ±band around historical HPP for the "Wajar" verdict
# Predicting landed HPP from a raw RMB price: the FINAL landed IDR per 1 RMB
# (already includes the forwarder's = Ocistok/Martkita margin + shipping + import),
# CALIBRATED from history as realized `HPP per Buah (Rp)` ÷ the `(x RMB)` note,
# per-SKU when it has ≥ MIN_LOTS Ocistok/Martkita lots, else the global median.
# It runs ~25% above the raw RMB→IDR spot (RMB_SPOT_FX_IDR) — that gap IS the
# forwarder margin + shipping + import. Fallback factor below.
RMB_TO_IDR_FALLBACK = 2832
RMB_SPOT_FX_IDR = 2250            # ≈ raw RMB→IDR spot, reference only (for the breakdown note)
RESTOCK_RMB_MIN_LOTS = 2
# Marketplace fee fallback (used only when sales history for a platform is too thin);
# the real fee = |admin| / omzet derived from BisaJual is preferred.
PLATFORM_FEE_FALLBACK = {"Shopee": 0.11, "Tokopedia": 0.10, "Tiktok": 0.25}
# restock_check.xlsx input columns
COL_RC_SKU = "SKU"
COL_RC_TOKO = "Toko"
COL_RC_RMB = "Harga RMB"          # raw supplier unit price in Yuan (optional)
COL_RC_HPP = "HPP IDR"           # landed cost per pc in Rupiah (optional; overrides RMB)
COL_RC_KMIN = "Kompetitor Min"   # competitor selling price, low end (Rp)
COL_RC_KMAX = "Kompetitor Max"   # competitor selling price, high end (Rp)
COL_RC_NOTE = "Catatan"

# --- Cash-flow restock plan (--cashflow) ---
# Turn the reorder metrics into a purchasing budget calendar: for each SKU that
# crosses its reorder point within the horizon, project WHEN to order, HOW MUCH,
# the COST (qty × replacement HPP = latest overseas lot price, fallback HPP_WA),
# and WHICH supplier — then bucket the Rupiah by month and supplier. Zero-config
# (built entirely from the stok/jual data), so it always runs in --all.
CASHFLOW_OUTPUT_FILENAME = "Analisa_Cashflow_Restock.xlsx"
CASHFLOW_HORIZON_MONTHS = 6       # plan window: only orders due within N months are budgeted
CASHFLOW_MAX_CYCLES = 60          # safety cap on reorder cycles simulated per SKU in the window

# --- Bundle / cross-sell market basket (--bundle) ---
# Find SKU pairs frequently bought in the same order (Invoice) for bundles /
# "frequently bought together". Zero-config (built from BisaJual).
BASKET_OUTPUT_FILENAME = "Analisa_Bundle_CrossSell.xlsx"
BASKET_MIN_PAIR_SUPPORT = 5      # min co-occurring orders before a SKU pair is reported
BASKET_TOP_N = 150               # cap on pair rows in the report

# --- Dead-stock / capital-release (--deadstock) ---
# Quantify capital frozen in Overstock + Slow/Dead SKUs and recommend how to free it
# (markdown / liquidate / stop-reorder). Built from the reorder metrics. Zero-config.
DEADSTOCK_OUTPUT_FILENAME = "Analisa_Modal_Beku.xlsx"
DEADSTOCK_DEAD_DAYS = 180        # no sale in N days → liquidate (vs markdown to speed turnover)

# --- Momentum + ABC focus (--momentum) ---
# Classify SKUs by sales momentum (recent vs prior window) and ABC (Pareto by trailing
# profit), to guide push vs prune. Built from BisaJual + HPP. Zero-config.
MOMENTUM_OUTPUT_FILENAME = "Analisa_Momentum_ABC.xlsx"
MOMENTUM_WINDOW_DAYS = 90         # each comparison window (recent vs the prior one)
MOMENTUM_MIN_QTY = 10             # min (recent + prior) qty before a SKU is classified
MOMENTUM_GROWTH_THRESHOLD = 0.30  # ± this vs the prior window → accelerating / declining
MOMENTUM_TRAILING_DAYS = 365      # window used to rank trailing profit for ABC
ABC_A_SHARE = 0.80                # cumulative trailing-profit share cutoff for class A
ABC_B_SHARE = 0.95               # ...and class B (the rest are class C)

# --- Sales trend & seasonality (--trend) ---
# Cross-year view: monthly/yearly omzet+profit time series, YoY growth, and a
# seasonal index (which calendar months consistently over/under-perform). Zero-config.
TREND_OUTPUT_FILENAME = "Analisa_Tren_Musiman.xlsx"
TREND_SEASONAL_MIN_YEARS = 2      # min complete years contributing before a month's seasonal index is shown
TREND_PEAK_INDEX = 1.15           # seasonal index ≥ this = peak month
TREND_LOW_INDEX = 0.85            # seasonal index ≤ this = soft month

# Source column names (raw from Excel with embedded newlines)
COL_STOK_SKU = "SKU"
COL_STOK_QTY = "Banyak\nBarang\n(Buah)"
COL_STOK_TOTAL_HPP = "Total\nHPP\n(Rp)"
COL_STOK_TANGGAL_BAYAR = "Tanggal\nBayar"
COL_STOK_TOKO = "Toko"                         # standardized supplier/forwarder column
COL_STOK_TOKO_LEGACY = "Toko[spasi]Akun Pemesan"   # pre-standardization header (auto-renamed on load)
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
