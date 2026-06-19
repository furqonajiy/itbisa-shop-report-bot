import logging

import pandas as pd

from utility.generic import build_report_path

from bisainvoice.generic import bisainvoice_to_excel


def generate_bisainvoice(ttk_file, df):
    logging.info("Generate BisaInvoice from {0} ({1} rows)".format(ttk_file, len(df)))

    # Remove Duplicate and Null Order ID
    df = df.dropna(axis=0, subset=['Order ID'])
    df = df.drop_duplicates(subset=['Order ID'])

    # Select Marketplace Tiktok / Tokopedia
    df['Marketplace'] = df['Purchase Channel'].str.title()

    # Select Needed Column
    df = df[['Created Time', 'Marketplace', 'Order ID',
             'Original Shipping Fee', 'Shipping Insurance']]

    # Change Column Name
    df.columns = ['Tanggal', 'Marketplace', 'Invoice', 'Ongkir', 'Asuransi']

    # Convert Data Type
    df['Tanggal'] = pd.to_datetime(df['Tanggal'], format='%d/%m/%Y %H:%M:%S').dt.strftime('%Y-%m-%d %H:%M:%S')  # Datetime
    df['Ongkir'] = df['Ongkir'].astype(str)
    df['Asuransi'] = df['Asuransi'].astype(str)

    # Export
    path = (ttk_file
            .replace(' v1', '')
            .replace(' v2', '')
            .replace('BisaTransaksi', 'BisaLaporan'))
    path = build_report_path(path)
    bisainvoice_to_excel(df, path, 'BisaInvoice Tiktok')
