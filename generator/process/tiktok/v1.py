import logging

import pandas as pd

from bisabonus.tiktok.v1 import generate_bisabonus
from bisainvoice.tiktok.v1 import generate_bisainvoice
from bisajual.tiktok.v1 import generate_bisajual
from bisaremit.tiktok.v1 import generate_bisaremit
from keywordchecker.tiktok import check_saldo_keyword, check_status_keyword
from utility.constant import BISALAPORAN_TIKTOK_DIR
from utility.generic import create_directory


def process(list_report):
    logging.info("Process Tiktok v2 File")

    create_directory(BISALAPORAN_TIKTOK_DIR)

    for tkp_file in list_report:
        read_bisatransaksi(tkp_file)

    for tkp_file in list_report:
        read_bisasaldo(tkp_file)


def read_bisatransaksi(tkp_file):
    cond1 = 'BisaTransaksi v2 Tiktok' in tkp_file
    cond2 = '~' not in tkp_file
    if cond1 and cond2:
        logging.debug("Read {0}".format(tkp_file))

        df = pd.read_excel(tkp_file, skiprows=4)

        # Remove rows with invalid status
        search_values = ['Dibatalkan']
        df = df[~df['Status Terakhir'].str.contains('|'.join(search_values))]

        if len(df) > 0:
            check_status_keyword("2", tkp_file, df)
            generate_bisainvoice(tkp_file, df)
            generate_bisajual(tkp_file, df)


def read_bisasaldo(tkp_file):
    cond1 = 'BisaSaldo v2 Tiktok' in tkp_file
    cond2 = '~' not in tkp_file
    if cond1 and cond2:
        logging.debug("Read {0}".format(tkp_file))

        df = pd.read_excel(tkp_file, skiprows=6)

        if len(df) > 0:
            check_saldo_keyword(tkp_file, df)
            generate_bisaremit(tkp_file, df)
            generate_bisabonus(tkp_file, df)
