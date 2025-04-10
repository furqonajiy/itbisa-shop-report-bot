import logging

import pandas as pd

from bisajual.generic import bisajual_to_excel
from utility.sku import standardize_sku


def generate_bisajual(tkp_file, df):
    logging.info("Generate BisaJual from {0} ({1} rows)".format(tkp_file, len(df)))

    # Assign Omzet as Jumlah Produk Dibeli multiply with Price
    df['Omzet'] = (df['Jumlah Produk Dibeli'] * df['Harga Jual (IDR)'].astype(str)
                   .str.replace('Rp ', '')
                   .str.replace('.', '')
                   .astype(int))

    # Assign 1PCS to SKU
    df['Pengali'] = df['Nomor SKU'].str.extract('(\d+)(?=\s*PCS)').apply(pd.to_numeric)
    df.loc[df['Pengali'].isnull(), 'Pengali'] = 1

    # Multiple Jumlah Produk Dibeli with PCS in SKU
    df['Jumlah Produk Dibeli'] = df['Jumlah Produk Dibeli'] * df['Pengali'].astype(int)

    # Remove XX PCS in SKU
    df['Nomor SKU'] = df['Nomor SKU'].str.replace('(\d+)PCS-', '', regex=True)

    # Select Needed Column
    df = df[['Nomor SKU', 'Nomor Invoice', 'Jumlah Produk Dibeli', 'Omzet']]

    # Change Column Name
    df.columns = ['SKU', 'Invoice', 'Kuantitas', 'Omzet']

    # Change to New SKU
    standardize_sku(df)

    # Export to Existing WorkBook
    path = (tkp_file
            .replace(' v1', '')
            .replace(' v2', '')
            .replace('BisaTransaksi', 'BisaLaporan'))
    bisajual_to_excel(df, path, 'BisaJual Tiktok')
