import logging

from keywordchecker.generic import handle_invalid_keywords

VALID_SALDO_KEYWORD = [
    # Not Used

    # Need to check

    # Nominal Remit

    # Keuntungan Tambahan

    # Kerugian Tambahan

    # BisaBonus
]

VALID_TRANSAKSI_KEYWORD = [
    # V1

    # V2
]

VALID_NOMINAL_REMIT_KEYWORD = [
    # Nominal Remit
    'Transaksi Penjualan Berhasil',
    'Pemotongan Ongkir',
    'Pemotongan untuk Asuransi',
]

VALID_KEUNTUNGAN_TAMBAHAN_KEYWORD = [
    # Keuntungan Tambahan
    'Subsidi Kupon Toko',
]

VALID_KERUGIAN_TAMBAHAN_KEYWORD = [
    # Kerugian Tambahan
    'Pemotongan Biaya Layanan',
    'Pemotongan Voucher Merchant',
    'Pemotongan Saldo untuk Kupon Toko',
    'Pemotongan Selisih Ongkir',
]

VALID_BONUS_KEYWORD = [
    # BisaBonus
    'PENAMBAHAN WALLET'
]


def check_saldo_keyword(tkp_file, df):
    logging.debug("Check BisaSaldo Keyword in {0}".format(tkp_file))

    invalid_rows = df[~df['Description'].str.contains('|'.join(VALID_SALDO_KEYWORD))]

    handle_invalid_keywords('BisaSaldo', tkp_file, invalid_rows)


def check_status_keyword(version, tkp_file, df):
    logging.debug("Check BisaTransaksi Keyword in {0}".format(tkp_file))

    if version == "1":
        invalid_rows = df[~df['Order Status'].isin(VALID_TRANSAKSI_KEYWORD)]
    elif version == "2":
        invalid_rows = df[~df['Status Terakhir'].isin(VALID_TRANSAKSI_KEYWORD)]

    handle_invalid_keywords('BisaTransaksi', tkp_file, invalid_rows)
