import logging

import pandas as pd

from laporan.bonus.shopee.v3 import generate_bonus
from laporan.fee.shopee.v3 import generate_fee
from laporan.invoice.shopee.v3 import generate_invoice
from laporan.jual.shopee.v3 import generate_jual
from laporan.remit.shopee.v3 import generate_remit
from laporan.keywordchecker.shopee import check_saldo_keyword, check_status_keyword


def process(list_report):
    logging.info("Process Shopee v3 File")

    for shp_file in list_report:
        read_transaksi(shp_file)

    df_fee = pd.DataFrame(columns=['Invoice', 'Potongan Pembayaran (Fee)', 'Nominal Remit (Fee)',
                                   'Keuntungan Tambahan (Fee)', 'Kerugian Tambahan (Fee)', 'Biaya Proses Pesanan'])
    for shp_file in list_report:
        df_fee = read_fee(shp_file, df_fee)

    for shp_file in list_report:
        read_saldo(shp_file, df_fee)


def read_transaksi(shp_file):
    cond1 = 'Transaksi v3 Shopee' in shp_file
    cond2 = '~' not in shp_file
    if cond1 and cond2:
        logging.debug("Read Transaksi Shopee {0}".format(shp_file))

        df = pd.read_excel(shp_file, dtype={
            'Harga Setelah Diskon': str,
            'Ongkir': str,
            'Alasan Pembatalan': str,
            'Jumlah': int,
            'Returned quantity': int
        })

        # Remove rows with invalid status
        search_values = ['Belum Bayar', 'Batal', 'Dibatalkan', 'Selesai']
        df = df[~df['Status Pesanan'].str.contains('|'.join(search_values))
                | df['Alasan Pembatalan'].str.contains('Paket hilang', na=False)
                | (df['Status Pesanan'].str.contains('Selesai') & (df['Jumlah'] != df['Returned quantity']))]

        if len(df) > 0:
            check_status_keyword("3", shp_file, df)
            generate_invoice(shp_file, df)
            generate_jual(shp_file, df)


def read_saldo(shp_saldo_file, df_fee):
    cond1 = 'Saldo v3 Shopee' in shp_saldo_file
    cond2 = '~' not in shp_saldo_file
    if cond1 and cond2:
        logging.debug("Read {0}".format(shp_saldo_file))

        df = pd.read_excel(shp_saldo_file, skiprows=17, dtype={'Jumlah': float, 'Saldo Akhir': float})

        if len(df) > 0:
            check_saldo_keyword(shp_saldo_file, df)
            generate_remit(shp_saldo_file, df, df_fee)
            generate_bonus(shp_saldo_file, df)


def _read_adjustment(shp_file):
    """Read the Shopee 'Adjustment' detail table.

    Its header row floats -- the preamble length depends on how many adjustment
    categories are summarised above the detail table -- so locate the
    'No. Pesanan Terhubung' header dynamically instead of a fixed skiprows.
    Returns [] when there is no Adjustment sheet or no detail table.
    """
    try:
        raw = pd.read_excel(shp_file, sheet_name='Adjustment', header=None, dtype=str)
    except ValueError as error:
        logging.debug(error)
        return []

    for i in range(len(raw)):
        if 'No. Pesanan Terhubung' in [str(c).strip() for c in raw.iloc[i].tolist()]:
            return pd.read_excel(shp_file, sheet_name='Adjustment', skiprows=i,
                                 dtype={'Biaya Penyesuaian': float})
    return []


def read_fee(shp_file, df_fee):
    cond1 = 'Fee v3 Shopee' in shp_file
    cond2 = '~' not in shp_file
    if cond1 and cond2:
        logging.debug("Read {0}".format(shp_file))

        df_raw = pd.read_excel(shp_file, sheet_name='Income', skiprows=5, dtype={
            'Harga Asli Produk': int, 'Total Diskon Produk': int, 'Biaya Administrasi': int, 'Biaya Layanan': int,
            'Biaya Proses Pesanan': int, 'Ongkir yang Diteruskan oleh Shopee ke Jasa Kirim': int, 'Total Penghasilan': int})

        df_adjust = _read_adjustment(shp_file)

        if len(df_raw) > 0:
            clean_df_fee = generate_fee(shp_file, df_raw, df_adjust)
            df_fee = pd.concat([df_fee, clean_df_fee])

    return df_fee
