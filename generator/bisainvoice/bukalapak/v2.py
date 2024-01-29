import logging

import pandas as pd

from bisainvoice.generic import bisainvoice_to_excel


def generate_bisainvoice(bl_file, df):
    logging.info("Generate BisaInvoice from {0} ({1} rows)".format(bl_file, len(df)))

    # Remove Duplicate
    df = df.drop_duplicates(subset=['ID Transaksi'])

    # Add New Column contain 'Bukalapak'
    df['Marketplace'] = 'Bukalapak'

    # Select Needed Column
    df = df[['Tanggal', 'Marketplace', 'ID Transaksi', 'Biaya Pengiriman', 'Biaya Asuransi']]

    # Change Column Name
    df.columns = ['Tanggal', 'Marketplace', 'Invoice', 'Ongkir', 'Asuransi']

    # Convert Data Type
    df['Tanggal'] = pd.to_datetime(df['Tanggal'], format='%Y-%m-%d %H:%M').dt.strftime('%Y-%m-%d %H:%M:%S')  # Datetime

    # Export
    path = bl_file.replace('BisaTransaksi', 'BisaLaporan')
    bisainvoice_to_excel(df, path, 'BisaInvoice Bukalapak')
