import logging

import pandas as pd

from bisainvoice.generic import bisainvoice_to_excel


def generate_bisainvoice(tkp_file, df):
    logging.info("Generate BisaInvoice from {0} ({1} rows)".format(tkp_file, len(df)))

    # Remove Duplicate and Null Nomor Invoice
    df = df.dropna(axis=0, subset=['Invoice'])
    df = df.drop_duplicates(subset=['Invoice'])

    # Add New Column contain 'Tokopedia'
    df['Marketplace'] = 'Tokopedia'

    # Select Needed Column
    df = df[['Payment Date', 'Marketplace', 'Invoice',
             'Shipping Price + fee (Rp.)', 'Insurance (Rp.)']]

    # Change Column Name
    df.columns = ['Tanggal', 'Marketplace', 'Invoice', 'Ongkir', 'Asuransi']

    # Convert Data Type
    df['Tanggal'] = pd.to_datetime(df['Tanggal'], format='%d-%m-%Y %H:%M:%S').dt.strftime('%Y-%m-%d %H:%M:%S')  # Datetime
    df['Ongkir'] = (df['Ongkir'].astype(str)
                    .str.replace('Rp ', '')
                    .str.replace('.', ''))
    df['Asuransi'] = (df['Asuransi'].astype(str)
                      .str.replace('Rp ', '')
                      .str.replace('.', ''))

    # Export
    path = (tkp_file
            .replace(' v1', '')
            .replace(' v2', '')
            .replace('BisaTransaksi', 'BisaLaporan'))
    bisainvoice_to_excel(df, path, 'BisaInvoice Tokopedia')
