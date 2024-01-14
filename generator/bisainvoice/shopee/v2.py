import logging

import pandas as pd

from bisainvoice.generic import bisainvoice_to_excel


def generate_bisainvoice(shp_file, df):
    logging.info("Generate BisaInvoice from {0} ({1} rows)".format(shp_file, len(df)))

    # Remove Duplicate and Null No. Pesanan
    df = df.dropna(axis=0, subset=['No. Pesanan'])
    df = df.drop_duplicates(subset=['No. Pesanan'])

    # Add New Column contain 'Shopee' and Asuransi
    df['Marketplace'] = 'Shopee'
    df['Asuransi'] = 0

    # Select Needed Column
    df = df[['Waktu Pembayaran Dilakukan', 'Marketplace', 'No. Pesanan',
             'Perkiraan Ongkos Kirim', 'Asuransi']]

    # Change Column Name
    df.columns = ['Tanggal', 'Marketplace', 'Invoice', 'Ongkir', 'Asuransi']

    # Convert Data Type
    df['Tanggal'] = pd.to_datetime(df['Tanggal'], format='%Y-%m-%d %H:%M').dt.strftime('%Y-%m-%d %H:%M:%S')  # Datetime

    df['Ongkir'] = df['Ongkir'].astype(float) * 1000
    df['Ongkir'] = df['Ongkir'].astype(int)

    df['Asuransi'] = (df['Asuransi'].astype(str)
                      .str.replace('Rp ', '')
                      .str.replace('.', ''))

    # Export
    path = (shp_file
            .replace(' v1', '')
            .replace(' v2', '')
            .replace('BisaTransaksi', 'BisaLaporan'))
    bisainvoice_to_excel(df, path, 'BisaInvoice Shopee')
