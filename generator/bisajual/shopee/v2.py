import logging

import pandas as pd

from bisajual.generic import bisajual_to_excel
from utility.sku import standardize_sku


def generate_bisajual(shp_file, df):
    logging.info("Generate BisaJual from {0} ({1} rows)".format(shp_file, len(df)))

    # Assign Omzet as Jumlah multiply with Price
    df['Omzet'] = (df['Jumlah'] * df['Harga Setelah Diskon'].astype(str)
                   .str.replace('Rp ', '')
                   .str.replace('.', '')
                   .astype(int))

    # Assign 1PCS to SKU
    df['Pengali'] = df['Nomor Referensi SKU'].str.extract('(\d+)(?=\s*PCS)').apply(pd.to_numeric)
    df.loc[df['Pengali'].isnull(), 'Pengali'] = 1

    # Multiple Jumlah with PCS in SKU
    df['Jumlah'] = df['Jumlah'] * df['Pengali'].astype(int)

    # Remove XX PCS in SKU
    df['Nomor Referensi SKU'] = df['Nomor Referensi SKU'].str.replace('(\d+)PCS-', '', regex=True)

    # Select Needed Column
    df = df[['Nomor Referensi SKU', 'No. Pesanan', 'Jumlah', 'Omzet']]

    # Change Column Name
    df.columns = ['SKU', 'Invoice', 'Kuantitas', 'Omzet']

    # Change to New SKU
    standardize_sku(df)

    # Export to Existing WorkBook
    path = (shp_file
            .replace(' v1', '')
            .replace(' v2', '')
            .replace('BisaTransaksi', 'BisaLaporan'))
    bisajual_to_excel(df, path, 'BisaJual Shopee')
