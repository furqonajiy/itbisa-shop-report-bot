import logging

import pandas as pd

from bisainvoice.generic import bisainvoice_to_excel


def generate_bisainvoice(tkp_file, df):
    logging.info("Generate BisaInvoice from {0} ({1} rows)".format(tkp_file, len(df)))

    # Remove Duplicate and Null Nomor Invoice
    df = df.dropna(axis=0, subset=['Nomor Invoice'])
    df = df.drop_duplicates(subset=['Nomor Invoice'])

    # Add New Column contain 'Tiktok'
    df['Marketplace'] = 'Tiktok'

    # Set Biaya Non Tunai to 0
    df.loc[(df['Biaya Pengiriman Tunai (IDR)'] == 'Non Tunai'), 'Biaya Pengiriman Tunai (IDR)'] = 0
    df.loc[(df['Biaya Asuransi Pengiriman (IDR)'] == 'Non Tunai'), 'Biaya Asuransi Pengiriman (IDR)'] = 0

    # Select Needed Column
    df = df[['Tanggal Pembayaran', 'Marketplace', 'Nomor Invoice',
             'Biaya Pengiriman Tunai (IDR)', 'Biaya Asuransi Pengiriman (IDR)']]

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
    bisainvoice_to_excel(df, path, 'BisaInvoice Tiktok')
