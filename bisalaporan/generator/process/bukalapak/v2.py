import logging

import pandas as pd

from invoice.bukalapak.v2 import generate_invoice
from jual.bukalapak.v2 import generate_jual
from remit.bukalapak.v2 import generate_remit
from keywordchecker.bukalapak import check_saldo_keyword, check_status_keyword


def process(list_report):
    logging.info("Process Bukalapak File")

    for bl_file in list_report:
        read_transaksi(bl_file)

    for bl_file in list_report:
        read_saldo(bl_file)


def read_transaksi(bl_file):
    cond1 = 'Transaksi v2 Bukalapak' in bl_file
    cond2 = '~' not in bl_file
    if cond1 and cond2:
        logging.debug("Read {0}".format(bl_file))

        df = pd.read_excel(bl_file, dtype={'Biaya Pengiriman': int, 'Biaya Asuransi': int})

        # Remove rows with invalid status
        search_values = ['Dikembalikan']
        df = df[~df['Status'].str.contains('|'.join(search_values))]

        if len(df) > 0:
            check_status_keyword(bl_file, df)
            generate_invoice(bl_file, df)
            generate_jual(bl_file, df)


def read_saldo(bl_file):
    cond1 = 'Saldo v2 Bukalapak' in bl_file
    cond2 = '~' not in bl_file
    if cond1 and cond2:
        logging.debug("Read {0}".format(bl_file))

        df = pd.read_csv(bl_file, dtype={'Mutasi': int})

        if len(df) > 0:
            check_saldo_keyword(bl_file, df)
            generate_remit(bl_file, df)
