import logging

import pandas as pd

from bisabonus.tokopedia import generate_bisabonus
from bisainvoice.tokopedia.v2 import generate_bisainvoice
from bisajual.tokopedia import generate_bisajual
from bisaremit.tokopedia import generate_bisaremit
from keywordchecker.tokopedia import check_saldo_keyword, check_status_keyword
from utility.constant import BISALAPORAN_TOKOPEDIA_DIR
from utility.generic import create_directory


def process(list_report):
    logging.info("Process Tokopedia File")

    create_directory(BISALAPORAN_TOKOPEDIA_DIR)

    for tkp_file in list_report:
        read_bisatransaksi(tkp_file)

    for tkp_file in list_report:
        read_bisasaldo(tkp_file)


def read_bisatransaksi(tkp_file):
    cond1 = 'BisaTransaksi v2 Tokopedia' in tkp_file
    cond2 = '~' not in tkp_file
    if cond1 and cond2:
        logging.debug("Read {0}".format(tkp_file))

        df = pd.read_excel(tkp_file, skiprows=4)

        # Remove rows with invalid status
        search_values = ['Dibatalkan']
        df = df[~df['Status Terakhir'].str.contains('|'.join(search_values))]

        if len(df) > 0:
            check_status_keyword(tkp_file, df)
            generate_bisainvoice(tkp_file, df)
            generate_bisajual(tkp_file, df)


def read_bisasaldo(tkp_file):
    cond1 = 'BisaSaldo v2 Tokopedia' in tkp_file
    cond2 = '~' not in tkp_file
    if cond1 and cond2:
        logging.debug("Read {0}".format(tkp_file))

        df = pd.read_excel(tkp_file, skiprows=6)

        if len(df) > 0:
            check_saldo_keyword(tkp_file, df)
            generate_bisaremit(tkp_file, df)
            generate_bisabonus(tkp_file, df)
