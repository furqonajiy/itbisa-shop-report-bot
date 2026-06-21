import io
import logging

import openpyxl
import pandas as pd

from laporan.invoice.tiktok.v1 import generate_invoice
from laporan.jual.tiktok.v1 import generate_jual
from laporan.remit.tiktok.v1 import generate_remit
from laporan.keywordchecker.tiktok import check_status_keyword


# TikTok exports the settlement "Fee" report in either English or the
# Indonesia-localized layout. Map the Indonesia headers the remit math needs
# back to the canonical English names so both versions process identically.
FEE_ID_TO_EN = {
    'ID Pesanan/Penyesuaian': 'Order/adjustment ID',
    'Waktu pemesanan': 'Order created time',
    'Jumlah penyelesaian pembayaran': 'Total settlement amount',
    'Total Pendapatan': 'Total Revenue',
    'Subtotal pengembalian dana setelah diskon penjual': 'Refund subtotal after seller discounts',
    'Ongkir yang ditalangi penyedia jasa logistik': 'Shipping costs passed on to the logistics provider',
}

# Read options for the TikTok order ("OrderSKUList") export. Row 2 is a column
# description that is skipped; IDs/strings stay text, quantities/prices stay int.
_ORDERS_READ_KW = dict(
    skiprows=lambda x: x == 1,
    dtype={'Order ID': str, 'Order Status': str,
           'Quantity': int, 'SKU Unit Original Price': int,
           'Original Shipping Fee': int, 'Shipping Insurance': int})


def process(list_report):
    logging.info("Process Tiktok v1 File")

    for ttk_file in list_report:
        read_transaksi(ttk_file)

    for ttk_file in list_report:
        read_fee(ttk_file)


def _read_orders(ttk_file):
    """Read a TikTok order export, resilient to a broken stored sheet dimension.

    Some TikTok exports save a wrong `<dimension>` in the .xlsx, so pandas/openpyxl
    read only the first column. When that happens (the 'Order Status' column is
    missing) re-open the workbook non-read-only -- which recomputes the real
    dimension -- and re-read identically.
    """
    df = pd.read_excel(ttk_file, **_ORDERS_READ_KW)
    if 'Order Status' not in df.columns:
        logging.debug("Sheet dimension looks wrong; re-reading %s", ttk_file)
        wb = openpyxl.load_workbook(ttk_file)
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        df = pd.read_excel(buf, **_ORDERS_READ_KW)
    return df


def read_transaksi(ttk_file):
    cond1 = 'Transaksi v1 Tiktok' in ttk_file
    cond2 = '~' not in ttk_file
    if cond1 and cond2:
        logging.debug("Read {0}".format(ttk_file))

        df = _read_orders(ttk_file)

        # Remove rows with invalid status
        df = df[(df['Tokopedia Invoice Number'].isna()) &
                (df['Order ID'] != 'Platform unique order ID.')]
        search_values = ['Belum dibayar', 'Dibatalkan']
        df = df[~df['Order Status'].str.contains('|'.join(search_values))]

        if len(df) > 0:
            check_status_keyword("1", ttk_file, df)
            generate_invoice(ttk_file, df)
            generate_jual(ttk_file, df)


def read_fee(ttk_file):
    cond1 = 'Fee v1 Tiktok' in ttk_file
    cond2 = '~' not in ttk_file
    if cond1 and cond2:
        logging.debug("Read {0}".format(ttk_file))

        # Read language-agnostically (English or Indonesia layout), then map the
        # Indonesia headers to the canonical English names the remit math uses.
        df = pd.read_excel(ttk_file, dtype=str)
        df = df.rename(columns=FEE_ID_TO_EN)
        # TikTok also varies the ID header casing/spacing ('Order/adjustment ID  ',
        # 'Order/Adjustment ID') -> normalize either to one canonical name.
        df = df.rename(columns=lambda c: 'Order/adjustment ID'
                       if str(c).strip().lower() == 'order/adjustment id' else c)

        # The remit math needs these as integer rupiah.
        for col in ('Total Revenue', 'Total settlement amount',
                    'Refund subtotal after seller discounts',
                    'Shipping costs passed on to the logistics provider'):
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

        if len(df) > 0:
            generate_remit(ttk_file, df)
