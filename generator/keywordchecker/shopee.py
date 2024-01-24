import logging

from keywordchecker.generic import handle_invalid_keywords

VALID_SALDO_KEYWORD = [
    # Not Used
    'Penarikan Dana',
    'Penggantian Dana Penuh',
    'Kompensasi kehilangan',
    'Shopee Ongkir',
    'Pengembalian Dana untuk Penarikan Gagal',

    # Need to check

    # Nominal Remit
    'Penghasilan dari Pesanan',

    # Keuntungan Tambahan

    # Kerugian Tambahan

    # BisaBonus
    'Cashback JNE',
]

VALID_TRANSAKSI_KEYWORD = [
    # V1

    # V2
    'Selesai',
    'Perlu Dikirim',
    'Sedang Dikirim',
    'Batal'
]

VALID_NOMINAL_REMIT_KEYWORD = [
    # Nominal Remit
    'Penghasilan dari Pesanan',
]

VALID_KEUNTUNGAN_TAMBAHAN_KEYWORD = [
    # Keuntungan Tambahan
    'AAA'
]

VALID_KERUGIAN_TAMBAHAN_KEYWORD = [
    # Kerugian Tambahan
    'AAA'
]

VALID_BONUS_KEYWORD = [
    # BisaBonus
    'PENAMBAHAN WALLET'
]


def check_saldo_keyword(shp_file, df):
    logging.debug("Check BisaSaldo Keyword in {0}".format(shp_file))

    invalid_rows = df[~df['Deskripsi'].str.contains('|'.join(VALID_SALDO_KEYWORD))]

    handle_invalid_keywords('BisaSaldo', shp_file, invalid_rows)


def check_status_keyword(version, shp_file, df):
    logging.info("Check BisaTransaksi Keyword in {0}".format(shp_file))

    if version == "1":
        invalid_rows = df[~df['Order Status'].isin(VALID_TRANSAKSI_KEYWORD)]
    elif version == "2":
        invalid_rows = df[~df['Status Pesanan'].isin(VALID_TRANSAKSI_KEYWORD)]

    handle_invalid_keywords('BisaTransaksi', shp_file, invalid_rows)
