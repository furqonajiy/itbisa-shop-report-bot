import logging

import pandas as pd

from laporan.invoice.tiktok.v1 import generate_invoice
from laporan.jual.tiktok.v1 import generate_jual
from laporan.remit.tiktok.v1 import generate_remit
from laporan.keywordchecker.tiktok import check_status_keyword


def process(list_report):
    logging.info("Process Tiktok v1 File")

    for ttk_file in list_report:
        read_transaksi(ttk_file)

    for ttk_file in list_report:
        read_fee(ttk_file)


def read_transaksi(ttk_file):
    cond1 = 'Transaksi v1 Tiktok' in ttk_file
    cond2 = '~' not in ttk_file
    if cond1 and cond2:
        logging.debug("Read {0}".format(ttk_file))

        df = pd.read_excel(ttk_file, skiprows=lambda x: x == 1,
                           dtype={'Order ID': str, 'Order Status': str,
                                  'Quantity': int, 'SKU Unit Original Price': int,
                                  'Original Shipping Fee': int, 'Shipping Insurance': int})

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

        df = pd.read_excel(ttk_file,
                           dtype={'Order/adjustment ID  ': str,
                                  'Order/Adjustment ID': str,
                                  'Order created time': str,
                                  'Total Revenue': int,
                                  'Total settlement amount': int,
                                  'Refund subtotal after seller discounts': int,
                                  'Shipping costs passed on to the logistics provider': int})

        # TikTok renamed 'Order/adjustment ID  ' (trailing spaces, lowercase) ->
        # 'Order/Adjustment ID'; normalize either header to one canonical name.
        df = df.rename(columns=lambda c: 'Order/adjustment ID'
                       if str(c).strip().lower() == 'order/adjustment id' else c)

        if len(df) > 0:
            generate_remit(ttk_file, df)
