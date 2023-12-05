import logging

import pandas as pd

from bisainvoice.tokopedia.v1 import generate_bisainvoice
from bisajual.tokopedia.v1 import generate_bisajual
from keywordchecker.tokopedia import check_status_keyword
from utility.constant import BISALAPORAN_TOKOPEDIA_V1_DIR, BISALAPORAN_TOKOPEDIA_V2_DIR
from utility.generic import create_directory


def process(list_report):
    logging.info("Process Tokopedia File")

    create_directory(BISALAPORAN_TOKOPEDIA_V1_DIR)
    create_directory(BISALAPORAN_TOKOPEDIA_V2_DIR)

    for tkp_file in list_report:
        read_bisatransaksi(tkp_file)

    # for tkp_file in list_report:
    #     read_bisasaldo(tkp_file)


def read_bisatransaksi(tkp_file):
    cond1 = 'BisaTransaksi v1 Tokopedia' in tkp_file
    cond2 = '~' not in tkp_file
    if cond1 and cond2:
        logging.debug("Read {0}".format(tkp_file))

        df = pd.read_excel(tkp_file, skiprows=3)
        df['Invoice'].fillna(method='ffill', inplace=True)
        df['Order Status'].fillna(method='ffill', inplace=True)

        # Remove rows with invalid status
        search_values = ['Dibatalkan']
        df = df[~df['Order Status'].str.contains('|'.join(search_values))]

        if len(df) > 0:
            check_status_keyword("1", tkp_file, df)
            generate_bisainvoice(tkp_file, df)
            generate_bisajual(tkp_file, df)
