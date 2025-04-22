import logging

import pandas as pd

from bisajual.generic import bisajual_to_excel
from utility.sku import standardize_sku


def generate_bisajual(ttk_file, df):
    logging.info("Generate BisaJual from {0} ({1} rows)".format(ttk_file, len(df)))

    # Assign Omzet as Quantity multiply with SKU Unit Original Price
    df['Omzet'] = df['Quantity'] * df['SKU Unit Original Price']

    # Assign 1PCS to SKU
    df['Pengali'] = df['Seller SKU'].str.extract('(\d+)(?=\s*PCS)').apply(pd.to_numeric)
    df.loc[df['Pengali'].isnull(), 'Pengali'] = 1

    # Multiple Quantity with PCS in SKU
    df['Quantity'] = df['Quantity'] * df['Pengali'].astype(int)

    # Remove XX PCS in SKU
    df['Seller SKU'] = df['Seller SKU'].str.replace('(\d+)PCS-', '', regex=True)

    # Select Needed Column
    df = df[['Seller SKU', 'Order ID', 'Quantity', 'Omzet']]

    # Change Column Name
    df.columns = ['SKU', 'Invoice', 'Kuantitas', 'Omzet']

    # Change to New SKU
    standardize_sku(df)

    # Export to Existing WorkBook
    path = (ttk_file
            .replace(' v1', '')
            .replace(' v2', '')
            .replace('BisaTransaksi', 'BisaLaporan'))
    bisajual_to_excel(df, path, 'BisaJual Tiktok')
