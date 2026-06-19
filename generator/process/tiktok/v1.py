import logging

import pandas as pd

from bisainvoice.tiktok.v1 import generate_bisainvoice
from bisajual.tiktok.v1 import generate_bisajual
from bisaremit.tiktok.v1 import generate_bisaremit
from keywordchecker.tiktok import check_status_keyword
from utility.constant import BISALAPORAN_TIKTOK_DIR
from utility.generic import create_directory


def process(list_report):
    logging.info("Process Tiktok v1 File")

    create_directory(BISALAPORAN_TIKTOK_DIR)

    for ttk_file in list_report:
        read_bisatransaksi(ttk_file)

    for ttk_file in list_report:
        read_bisafee(ttk_file)


def read_bisatransaksi(ttk_file):
    cond1 = 'BisaTransaksi v1 Tiktok' in ttk_file
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
            generate_bisainvoice(ttk_file, df)
            generate_bisajual(ttk_file, df)


def read_bisafee(ttk_file):
    cond1 = 'BisaFee v1 Tiktok' in ttk_file
    cond2 = '~' not in ttk_file
    if cond1 and cond2:
        logging.debug("Read {0}".format(ttk_file))

        df = pd.read_excel(ttk_file,
                           dtype={'Order/adjustment ID  ': str,
                                  'Order created time': str,
                                  'Total Revenue': int,
                                  'Total settlement amount': int,
                                  'Refund subtotal after seller discounts': int,
                                  'Shipping costs passed on to the logistics provider': int})

        if len(df) > 0:
            generate_bisaremit(ttk_file, df)
