import logging

import pandas as pd

from bisainvoice.generic import bisainvoice_to_excel


def generate_bisainvoice(ttk_file, df):
    logging.info("Generate BisaInvoice from {0} ({1} rows)".format(ttk_file, len(df)))

    # Remove Duplicate and Null Order ID
    df = df.dropna(axis=0, subset=['Order ID'])
    df = df.drop_duplicates(subset=['Order ID'])

    # Add New Column contain 'Tiktok'
    df['Marketplace'] = 'Tiktok'

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
    bisainvoice_to_excel(df, path, 'BisaInvoice Tiktok')
