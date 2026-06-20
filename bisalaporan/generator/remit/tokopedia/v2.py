import logging

import pandas as pd

from utility.generic import build_report_path

from remit.generic import remit_to_excel
from keywordchecker.tokopedia import VALID_NOMINAL_REMIT_KEYWORD, VALID_KEUNTUNGAN_TAMBAHAN_KEYWORD, VALID_KERUGIAN_TAMBAHAN_KEYWORD, VALID_POTONGAN_PEMBAYARAN_KEYWORD


def generate_remit(tkp_file, df):
    logging.info("Generate Remit Tokopedia from {0} ({1} rows)".format(tkp_file, len(df)))

    # Select rows which contains invoice number
    df = df[df['Description'].str.contains('INV')]

    # Generate Invoice from Description
    df['Invoice'] = df['Description'].str.extract(r'(INV\S+)')

    # Initialize Biaya Layanan and Remit
    df['Nominal Remit'] = 0
    df['Potongan Pembayaran'] = 0
    df['Keuntungan Tambahan'] = 0
    df['Kerugian Tambahan'] = 0

    # Generate Nominal Remit, Keuntungan Tambahan, Kerugian Tambahan, based on Description
    df.loc[df['Description'].str.contains('|'.join(VALID_NOMINAL_REMIT_KEYWORD)), 'Nominal Remit'] = df['Nominal (Rp)']
    df.loc[df['Description'].str.contains('|'.join(VALID_POTONGAN_PEMBAYARAN_KEYWORD)), 'Potongan Pembayaran'] = -df['Nominal (Rp)']
    df.loc[df['Description'].str.contains('|'.join(VALID_KEUNTUNGAN_TAMBAHAN_KEYWORD)), 'Keuntungan Tambahan'] = df['Nominal (Rp)']
    df.loc[df['Description'].str.contains('|'.join(VALID_KERUGIAN_TAMBAHAN_KEYWORD)), 'Kerugian Tambahan'] = -df['Nominal (Rp)']

    # Select Needed Column
    df = df[['Invoice', 'Date', 'Potongan Pembayaran', 'Nominal Remit',
             'Keuntungan Tambahan', 'Kerugian Tambahan']]

    # Change Column Name
    df.columns = ['Invoice', 'Tanggal Remit', 'Potongan Pembayaran', 'Nominal Remit',
                  'Keuntungan Tambahan', 'Kerugian Tambahan']

    # Convert Data Type
    df['Tanggal Remit'] = pd.to_datetime(df['Tanggal Remit'], format='mixed', dayfirst=False).dt.strftime('%Y-%m-%d %H:%M')  # Datetime

    # Aggregate
    df = df.groupby(['Invoice', 'Tanggal Remit']).sum().sort_values('Invoice')

    # Export to Existing WorkBook
    path = (tkp_file
            .replace(' v1', '')
            .replace(' v2', '')
            .replace('Saldo', 'Laporan'))
    path = build_report_path(path)
    remit_to_excel(df, path, 'Remit Tokopedia')
