import logging

import pandas as pd

from bisaremit.generic import bisaremit_to_excel
from keywordchecker.shopee import VALID_NOMINAL_REMIT_KEYWORD, VALID_KEUNTUNGAN_TAMBAHAN_KEYWORD, VALID_KERUGIAN_TAMBAHAN_KEYWORD


def generate_bisaremit(shp_file, df, df_fee):
    logging.info("Generate BisaRemit Shopee from {0} ({1} rows)".format(shp_file, len(df)))

    # Select rows which contains invoice number
    df = df[df['Deskripsi'].str.contains('#')]

    # Generate Invoice from Deskripsi
    df['Invoice'] = df['Deskripsi'].str.replace('^.*(?=#)', '')
    df['Invoice'] = df['Invoice'].str.replace('#', '')

    # Initialize Biaya Layanan and Remit
    df['Potongan Pembayaran'] = 0
    df['Nominal Remit'] = 0
    df['Keuntungan Tambahan'] = 0
    df['Kerugian Tambahan'] = 0

    # Generate Nominal Remit, Keuntungan Tambahan, Kerugian Tambahan, based on Deskripsi
    df.loc[df['Deskripsi'].str.contains('|'.join(VALID_NOMINAL_REMIT_KEYWORD)), 'Nominal Remit'] = df['Jumlah Dana']

    df.loc[df['Deskripsi'].str.contains('|'.join(VALID_KEUNTUNGAN_TAMBAHAN_KEYWORD)), 'Keuntungan Tambahan'] = df['Jumlah Dana']

    df.loc[df['Deskripsi'].str.contains('|'.join(VALID_KERUGIAN_TAMBAHAN_KEYWORD)), 'Kerugian Tambahan'] = -df['Jumlah Dana']

    # Select Needed Column
    df = df[['Invoice', 'Tanggal', 'Potongan Pembayaran', 'Nominal Remit',
             'Keuntungan Tambahan', 'Kerugian Tambahan']]

    # Change Column Name
    df.columns = ['Invoice', 'Tanggal Remit', 'Potongan Pembayaran', 'Nominal Remit',
                  'Keuntungan Tambahan', 'Kerugian Tambahan']

    # Convert Data Type
    df['Tanggal Remit'] = pd.to_datetime(df['Tanggal Remit'], format='%Y-%m-%d %H:%M').dt.strftime('%Y-%m-%d %H:%M')  # Datetime

    # Aggregate
    df = df.groupby(['Invoice', 'Tanggal Remit']).sum().sort_values('Invoice').reset_index()

    # Left Join with BisaFee
    df = df.merge(df_fee, on=['Invoice', 'Nominal Remit'], how='left')
    df['Potongan Pembayaran'] = df['Potongan Pembayaran'] + df['Potongan Pembayaran (Fee)']
    df['Keuntungan Tambahan'] = df['Keuntungan Tambahan'] + df['Keuntungan Tambahan (Fee)']
    df['Kerugian Tambahan'] = df['Kerugian Tambahan'] + df['Kerugian Tambahan (Fee)']

    # Select Needed Column after Join
    df = df[['Invoice', 'Tanggal Remit', 'Potongan Pembayaran', 'Nominal Remit',
             'Keuntungan Tambahan', 'Kerugian Tambahan']].set_index('Invoice')

    # Export to Existing WorkBook
    path = (shp_file
            .replace(' v1', '')
            .replace(' v2', '')
            .replace('BisaSaldo', 'BisaLaporan')
            .replace('.csv', '.xlsx'))
    bisaremit_to_excel(df, path, 'BisaRemit Shopee')
