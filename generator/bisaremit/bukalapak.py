import logging

import pandas as pd

from bisaremit.generic import bisaremit_to_excel


def generate_bisaremit(bl_file, df):
    logging.info("Generate BisaRemit Bukalapak from {0} ({1} rows)".format(bl_file, len(df)))

    # Select rows which contains invoice number
    df = df[df['Keterangan'].str.contains('#')]

    # Generate Invoice from Deskripsi
    df['Invoice'] = df['Keterangan'].str.extract(r'#(\S+)')
    df['Invoice'] = df['Invoice'].str.replace('#', '')

    # Initialize Biaya Layanan and Remit
    df['Potongan Pembayaran'] = 0
    df['Nominal Remit'] = 0
    df['Keuntungan Tambahan'] = 0
    df['Kerugian Tambahan'] = 0

    # Generate Biaya Layanan and Remit based on Deskripsi
    df.loc[df['Keterangan'].str.contains('Remit untuk transaksi'), 'Nominal Remit'] = df['Mutasi']

    df.loc[df['Keterangan'].str.contains('Pemotongan biaya') |
           df['Keterangan'].str.contains('Pembayaran fee'), 'Kerugian Tambahan'] = -df['Mutasi']

    # Select Needed Column
    df = df[['Invoice', 'Waktu', 'Potongan Pembayaran', 'Nominal Remit',
             'Keuntungan Tambahan', 'Kerugian Tambahan']]

    # Change Column Name
    df.columns = ['Invoice', 'Tanggal Remit', 'Potongan Pembayaran', 'Nominal Remit',
                  'Keuntungan Tambahan', 'Kerugian Tambahan']

    # Convert Data Type
    df['Tanggal Remit'] = (pd.to_datetime(df['Tanggal Remit'], format='%d %B %Y %H:%M')
                           .dt.strftime('%Y-%m-%d %H:%M'))  # Datetime

    # Aggregate
    df = df.groupby(['Invoice', 'Tanggal Remit']).sum().sort_values('Invoice')

    # Export to Existing WorkBook
    path = bl_file.replace('BisaSaldo', 'BisaLaporan').replace('csv', 'xlsx')
    bisaremit_to_excel(df, path, 'BisaRemit Bukalapak')
