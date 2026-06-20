import logging

import pandas as pd

from bonus.tokopedia.v2 import generate_bonus
from invoice.tokopedia.v2 import generate_invoice
from jual.tokopedia.v2 import generate_jual
from remit.tokopedia.v2 import generate_remit
from keywordchecker.tokopedia import check_saldo_keyword, check_status_keyword


def process(list_report):
    logging.info("Process Tokopedia v2 File")

    for tkp_file in list_report:
        read_transaksi(tkp_file)

    for tkp_file in list_report:
        read_saldo(tkp_file)


def read_transaksi(tkp_file):
    cond1 = 'Transaksi v2 Tokopedia' in tkp_file
    cond2 = '~' not in tkp_file
    if cond1 and cond2:
        logging.debug("Read {0}".format(tkp_file))

        df = pd.read_excel(tkp_file, skiprows=4)

        # Remove rows with invalid status
        search_values = ['Dibatalkan']
        df = df[~df['Status Terakhir'].str.contains('|'.join(search_values))]

        if len(df) > 0:
            check_status_keyword("2", tkp_file, df)
            generate_invoice(tkp_file, df)
            generate_jual(tkp_file, df)


def read_saldo(tkp_file):
    cond1 = 'Saldo v2 Tokopedia' in tkp_file
    cond2 = '~' not in tkp_file
    if cond1 and cond2:
        logging.debug("Read {0}".format(tkp_file))

        df = pd.read_excel(tkp_file, skiprows=6)

        if len(df) > 0:
            check_saldo_keyword(tkp_file, df)
            generate_remit(tkp_file, df)
            generate_bonus(tkp_file, df)
