import logging

import pandas as pd

from laporan.invoice.tokopedia.v1 import generate_invoice
from laporan.jual.tokopedia.v1 import generate_jual
from laporan.keywordchecker.tokopedia import check_status_keyword


def process(list_report):
    logging.info("Process Tokopedia v1 File")

    for tkp_file in list_report:
        read_transaksi(tkp_file)


def read_transaksi(tkp_file):
    cond1 = 'Transaksi v1 Tokopedia' in tkp_file
    cond2 = '~' not in tkp_file
    if cond1 and cond2:
        logging.debug("Read {0}".format(tkp_file))

        df = pd.read_excel(tkp_file, skiprows=3)
        df['Invoice'] = df['Invoice'].ffill()
        df['Order Status'] = df['Order Status'].ffill()

        # Remove rows with invalid status
        search_values = ['Dibatalkan']
        df = df[~df['Order Status'].str.contains('|'.join(search_values))]

        if len(df) > 0:
            check_status_keyword("1", tkp_file, df)
            generate_invoice(tkp_file, df)
            generate_jual(tkp_file, df)
