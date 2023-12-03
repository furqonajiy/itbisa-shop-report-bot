import logging

import pandas as pd

from bisainvoice.bukalapak import generate_bisainvoice
from bisajual.bukalapak import generate_bisajual
from bisaremit.bukalapak import generate_bisaremit
from keywordchecker.bukalapak import check_saldo_keyword, check_status_keyword
from utility.constant import BISALAPORAN_BUKALAPAK_DIR
from utility.generic import create_directory


def process(list_report):
    logging.info("Process Bukalapak File")

    create_directory(BISALAPORAN_BUKALAPAK_DIR)

    for bl_file in list_report:
        read_bisatransaksi(bl_file)

    for bl_file in list_report:
        read_bisasaldo(bl_file)


def read_bisatransaksi(bl_file):
    cond1 = 'BisaTransaksi v2 Bukalapak' in bl_file
    cond2 = '~' not in bl_file
    if cond1 and cond2:
        logging.debug("Read {0}".format(bl_file))

        df = pd.read_excel(bl_file, dtype={'Biaya Pengiriman': int, 'Biaya Asuransi': int})

        # Remove rows with invalid status
        search_values = ['Dikembalikan']
        df = df[~df['Status'].str.contains('|'.join(search_values))]

        if len(df) > 0:
            check_status_keyword(bl_file, df)
            generate_bisainvoice(bl_file, df)
            generate_bisajual(bl_file, df)


def read_bisasaldo(bl_file):
    cond1 = 'BisaSaldo v2 Bukalapak' in bl_file
    cond2 = '~' not in bl_file
    if cond1 and cond2:
        logging.debug("Read {0}".format(bl_file))

        df = pd.read_csv(bl_file, dtype={'Mutasi': int})

        if len(df) > 0:
            check_saldo_keyword(bl_file, df)
            generate_bisaremit(bl_file, df)
