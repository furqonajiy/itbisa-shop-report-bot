import logging

import pandas as pd

from bisaremit.generic import bisaremit_to_excel
from keywordchecker.shopee import VALID_NOMINAL_REMIT_KEYWORD, VALID_KEUNTUNGAN_TAMBAHAN_KEYWORD, VALID_KERUGIAN_TAMBAHAN_KEYWORD


def generate_bisaremit(tkp_file, df):
    logging.info("Generate BisaRemit Shopee from {0} ({1} rows)".format(tkp_file, len(df)))

    # Select rows which contains invoice number
    df = df[df['Deskripsi'].str.contains('#')]

    # Generate Invoice from Description
    df['Invoice'] = df['Deskripsi'].str.replace('^.*(?=#)', '')
    df['Invoice'] = df['Invoice'].str.replace('#', '')

    # Initialize Biaya Layanan and Remit
    df['Potongan Pembayaran'] = 0
    df['Nominal Remit'] = 0
    df['Keuntungan Tambahan'] = 0
    df['Kerugian Tambahan'] = 0

    # Generate Nominal Remit, Keuntungan Tambahan, Kerugian Tambahan, based on Description
    df.loc[df['Description'].str.contains('|'.join(VALID_NOMINAL_REMIT_KEYWORD)), 'Nominal Remit'] = df['Nominal (Rp)']

    df.loc[df['Description'].str.contains('|'.join(VALID_KEUNTUNGAN_TAMBAHAN_KEYWORD)), 'Keuntungan Tambahan'] = df['Nominal (Rp)']

    df.loc[df['Description'].str.contains('|'.join(VALID_KERUGIAN_TAMBAHAN_KEYWORD)), 'Kerugian Tambahan'] = -df['Nominal (Rp)']

    # Select Needed Column
    df = df[['Invoice', 'Tanggal', 'Potongan Pembayaran', 'Nominal Remit',
             'Keuntungan Tambahan', 'Kerugian Tambahan']]

    # Change Column Name
    df.columns = ['Invoice', 'Tanggal Remit', 'Potongan Pembayaran', 'Nominal Remit',
                  'Keuntungan Tambahan', 'Kerugian Tambahan']

    # Convert Data Type
    df['Tanggal Remit'] = pd.to_datetime(df['Tanggal Remit'], format='%Y-%m-%d %H:%M').dt.strftime('%Y-%m-%d %H:%M')  # Datetime

    # Aggregate
    df = df.groupby(['Invoice', 'Tanggal Remit']).sum().sort_values('Invoice')

    # Export to Existing WorkBook
    path = (tkp_file
            .replace(' v1', '')
            .replace(' v2', '')
            .replace('BisaSaldo', 'BisaLaporan'))
    bisaremit_to_excel(df, path, 'BisaRemit Shopee')
