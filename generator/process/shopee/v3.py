import logging

import pandas as pd

from bisabonus.shopee.v3 import generate_bisabonus
from bisafee.shopee.v3 import generate_bisafee
from bisainvoice.shopee.v3 import generate_bisainvoice
from bisajual.shopee.v3 import generate_bisajual
from bisaremit.shopee.v3 import generate_bisaremit
from keywordchecker.shopee import check_saldo_keyword, check_status_keyword
from utility.constant import BISALAPORAN_SHOPEE_DIR
from utility.generic import create_directory


def process(list_report):
    logging.info("Process Shopee v3 File")

    create_directory(BISALAPORAN_SHOPEE_DIR)

    for shp_file in list_report:
        read_bisatransaksi(shp_file)

    df_fee = pd.DataFrame(columns=['Invoice', 'Potongan Pembayaran (Fee)', 'Nominal Remit (Fee)', 'Keuntungan Tambahan (Fee)', 'Kerugian Tambahan (Fee)'])
    for shp_file in list_report:
        df_fee = read_bisafee(shp_file, df_fee)

    for shp_file in list_report:
        read_bisasaldo(shp_file, df_fee)


def read_bisatransaksi(shp_file):
    cond1 = 'BisaTransaksi v3 Shopee' in shp_file
    cond2 = '~' not in shp_file
    if cond1 and cond2:
        logging.debug("Read BisaTransaksi Shopee {0}".format(shp_file))

        df = pd.read_excel(shp_file, dtype={'Harga Setelah Diskon': str, 'Ongkir': str, 'Alasan Pembatalan': str})

        # Remove rows with invalid status
        search_values = ['Belum Bayar', 'Batal', 'Dibatalkan']
        df = df[~df['Status Pesanan'].str.contains('|'.join(search_values)) | df['Alasan Pembatalan'].str.contains('Paket hilang', na=False)]

        if len(df) > 0:
            check_status_keyword("2", shp_file, df)
            generate_bisainvoice(shp_file, df)
            generate_bisajual(shp_file, df)


def read_bisasaldo(shp_file, df_fee):
    cond1 = 'BisaSaldo v3 Shopee' in shp_file
    cond2 = '~' not in shp_file
    if cond1 and cond2:
        logging.debug("Read {0}".format(shp_file))

        df = pd.read_excel(shp_file, skiprows=17, dtype={'Jumlah': float, 'Saldo Akhir': float})

        if len(df) > 0:
            check_saldo_keyword(shp_file, df)
            generate_bisaremit(shp_file, df, df_fee)
            generate_bisabonus(shp_file, df)


def read_bisafee(shp_file, df_fee):
    cond1 = 'BisaFee v3 Shopee' in shp_file
    cond2 = '~' not in shp_file
    if cond1 and cond2:
        logging.debug("Read {0}".format(shp_file))

        df_raw = pd.read_excel(shp_file, sheet_name='Income', skiprows=5, dtype={
            'Harga Asli Produk': int, 'Total Diskon Produk': int, 'Biaya Administrasi': int, 'Biaya Layanan (termasuk PPN 11%)': int,
            'Ongkir yang Diteruskan oleh Shopee ke Jasa Kirim': int, 'Total Penghasilan': int})

        df_adjust = []
        try:
            df_adjust = pd.read_excel(shp_file, sheet_name='Adjustment', skiprows=14, dtype={
                'Biaya Penyesuaian': float
            })
        except ValueError as error:
            logging.debug(error)

        if len(df_raw) > 0:
            clean_df_fee = generate_bisafee(shp_file, df_raw, df_adjust)
            df_fee = pd.concat([df_fee, clean_df_fee])

    return df_fee
