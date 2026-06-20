import logging

import pandas as pd

from bonus.shopee.v2 import generate_bonus
from fee.shopee.v2 import generate_fee
from invoice.shopee.v2 import generate_invoice
from jual.shopee.v2 import generate_jual
from remit.shopee.v2 import generate_remit
from keywordchecker.shopee import check_saldo_keyword, check_status_keyword


def process(list_report):
    logging.info("Process Shopee v2 File")

    for shp_file in list_report:
        read_transaksi(shp_file)

    df_fee = pd.DataFrame(columns=['Invoice', 'Potongan Pembayaran (Fee)', 'Nominal Remit (Fee)', 'Keuntungan Tambahan (Fee)', 'Kerugian Tambahan (Fee)'])
    for shp_file in list_report:
        df_fee = read_fee(shp_file, df_fee)

    for shp_file in list_report:
        read_saldo(shp_file, df_fee)


def read_transaksi(shp_file):
    cond1 = 'Transaksi v2 Shopee' in shp_file
    cond2 = '~' not in shp_file
    if cond1 and cond2:
        logging.debug("Read Transaksi Shopee {0}".format(shp_file))

        df = pd.read_excel(shp_file, dtype={'Harga Setelah Diskon': str, 'Ongkir': str, 'Alasan Pembatalan': str})

        # Remove rows with invalid status
        search_values = ['Batal', 'Dibatalkan']
        df = df[~df['Status Pesanan'].str.contains('|'.join(search_values)) | df['Alasan Pembatalan'].str.contains('Paket hilang', na=False)]

        if len(df) > 0:
            check_status_keyword("2", shp_file, df)
            generate_invoice(shp_file, df)
            generate_jual(shp_file, df)


def read_saldo(shp_file, df_fee):
    cond1 = 'Saldo v2 Shopee' in shp_file
    cond2 = '~' not in shp_file
    if cond1 and cond2:
        logging.debug("Read {0}".format(shp_file))

        df = pd.read_csv(shp_file, skiprows=6, dtype={'Jumlah Dana': int})

        if len(df) > 0:
            check_saldo_keyword(shp_file, df)
            generate_remit(shp_file, df, df_fee)
            generate_bonus(shp_file, df)


def read_fee(shp_file, df_fee):
    cond1 = 'Fee v2 Shopee' in shp_file
    cond2 = '~' not in shp_file
    if cond1 and cond2:
        logging.debug("Read {0}".format(shp_file))

        df_raw = pd.read_excel(shp_file, sheet_name='Income', skiprows=5, dtype={
            'Harga Asli Produk': int, 'Total Diskon Produk': int, 'Biaya Administrasi': int, 'Biaya Layanan (termasuk PPN 11%)': int,
            'Ongkir yang Diteruskan oleh Shopee ke Jasa Kirim': int, 'Total Penghasilan': int})

        df_adjust = []
        try:
            df_adjust = pd.read_excel(shp_file, sheet_name='Adjustment', skiprows=13, dtype={
                'Biaya Penyesuaian': float
            })
        except ValueError as error:
            logging.debug(error)

        if len(df_raw) > 0:
            clean_df_fee = generate_fee(shp_file, df_raw, df_adjust)
            df_fee = pd.concat([df_fee, clean_df_fee])

    return df_fee
