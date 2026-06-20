import logging

import pandas as pd

from utility.generic import build_report_path

from jual.generic import jual_to_excel
from utility.sku import standardize_sku


def generate_jual(bl_file, df):
    logging.info("Generate Jual from {0} ({1} rows)".format(bl_file, len(df)))

    # Assign 1PCS to SKU
    df['Pengali'] = df['SKU'].str.extract('(\d+)(?=\s*PCS)').apply(pd.to_numeric)
    df.loc[df['Pengali'].isnull(), 'Pengali'] = 1

    # Multiple Jumlah Produk with PCS in SKU
    df['Jumlah Produk'] = df['Jumlah Produk'] * df['Pengali'].astype(int)

    # Remove XX PCS in SKU
    df['SKU'] = df['SKU'].str.replace('(\d+)PCS-', '', regex=True)

    # Select Needed Column
    df = df[['SKU', 'ID Transaksi', 'Jumlah Produk', 'Harga Produk']]

    # Change Column Name
    df.columns = ['SKU', 'Invoice', 'Banyak', 'Omzet']

    # Change to New SKU
    standardize_sku(df)

    # Export to Existing WorkBook
    path = bl_file.replace('Transaksi', 'Laporan')
    path = build_report_path(path)
    jual_to_excel(df, path, 'Jual Bukalapak')
