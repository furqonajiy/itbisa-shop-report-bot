import logging

import pandas as pd

from laporan.utility.generic import build_report_path

from laporan.invoice.generic import invoice_to_excel


def generate_invoice(bl_file, df):
    logging.info("Generate Invoice from {0} ({1} rows)".format(bl_file, len(df)))

    # Remove Duplicate
    df = df.drop_duplicates(subset=['ID Transaksi'])

    # Add New Column contain 'Bukalapak'
    df['Marketplace'] = 'Bukalapak'

    # Select Needed Column
    df = df[['Tanggal', 'Marketplace', 'ID Transaksi', 'Biaya Pengiriman', 'Biaya Asuransi']]

    # Change Column Name
    df.columns = ['Tanggal', 'Marketplace', 'Invoice', 'Ongkir', 'Asuransi']

    # Convert Data Type
    df['Tanggal'] = pd.to_datetime(df['Tanggal'], format='mixed', dayfirst=False).dt.strftime('%Y-%m-%d %H:%M:%S')  # Datetime

    # Export
    path = bl_file.replace('Transaksi', 'Laporan')
    path = build_report_path(path)
    invoice_to_excel(df, path, 'Invoice Bukalapak')
