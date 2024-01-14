import logging

import pandas as pd

from bisaremit.generic import bisaremit_to_excel


def generate_bisaremit(tkp_file, df):
    logging.info("Generate BisaRemit Tokopedia from {0} ({1} rows)".format(tkp_file, len(df)))

    # Select rows which contains invoice number
    df = df[df['Description'].str.contains('INV')]

    # Generate Invoice from Description
    df['Invoice'] = df['Description'].str.extract(r'(INV\S+)')

    # Initialize Biaya Layanan and Remit
    df['Potongan Pembayaran'] = 0
    df['Nominal Remit'] = 0
    df['Keuntungan Tambahan'] = 0
    df['Kerugian Tambahan'] = 0

    # Generate Biaya Layanan and Remit based on Description
    df.loc[df['Description'].str.contains('Transaksi Penjualan Berhasil') |
           df['Description'].str.contains('Pemotongan Ongkir') |
           df['Description'].str.contains('Pemotongan untuk Asuransi'), 'Nominal Remit'] = df['Nominal (Rp)']

    df.loc[df['Description'].str.contains('Subsidi Kupon Toko') |
           df['Description'].str.contains('Pemotongan Biaya Layanan') |
           df['Description'].str.contains('Pemotongan Voucher Merchant') |
           df['Description'].str.contains('Pemotongan Saldo untuk Kupon Toko') |
           df['Description'].str.contains('Pemotongan Selisih Ongkir') |
           df['Description'].str.contains('Subsidi Kupon Toko'), 'Kerugian Tambahan'] = -df['Nominal (Rp)']

    # Select Needed Column
    df = df[['Invoice', 'Date', 'Potongan Pembayaran', 'Nominal Remit',
             'Keuntungan Tambahan', 'Kerugian Tambahan']]

    # Change Column Name
    df.columns = ['Invoice', 'Tanggal Remit', 'Potongan Pembayaran', 'Nominal Remit',
                  'Keuntungan Tambahan', 'Kerugian Tambahan']

    # Convert Data Type
    df['Tanggal Remit'] = (pd.to_datetime(df['Tanggal Remit'], format='%Y-%m-%d %H:%M')
                           .dt.strftime('%Y-%m-%d %H:%M'))  # Datetime

    # Aggregate
    df = df.groupby(['Invoice', 'Tanggal Remit']).sum().sort_values('Invoice')

    # Export to Existing WorkBook
    path = (tkp_file
            .replace(' v1', '')
            .replace(' v2', '')
            .replace('BisaSaldo', 'BisaLaporan'))
    bisaremit_to_excel(df, path, 'BisaRemit Tokopedia')
