import logging

import pandas as pd

from laporan.utility.generic import build_report_path

from laporan.remit.generic import remit_to_excel


def generate_remit(tkp_file, df):
    logging.info("Generate Remit Tiktok from {0} ({1} rows)".format(tkp_file, len(df)))

    # Generate Invoice from Order/adjustment ID
    df = df[['Order/adjustment ID', 'Order created time', 'Total settlement amount', 'Total Revenue',
             'Shipping costs passed on to the logistics provider', 'Refund subtotal after seller discounts']]

    # Initialize Biaya Layanan and Remit
    df['Nominal Remit'] = df['Total Revenue'] - df['Refund subtotal after seller discounts']
    df['Potongan Pembayaran'] = -df['Shipping costs passed on to the logistics provider']
    df['Keuntungan Tambahan'] = 0
    df['Kerugian Tambahan'] = df['Total Revenue'] - df['Total settlement amount'] - df['Refund subtotal after seller discounts']

    # Select Needed Column
    df = df[['Order/adjustment ID', 'Order created time', 'Potongan Pembayaran', 'Nominal Remit',
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
            .replace('Fee', 'Laporan'))
    path = build_report_path(path)
    remit_to_excel(df, path, 'Remit Tiktok')
