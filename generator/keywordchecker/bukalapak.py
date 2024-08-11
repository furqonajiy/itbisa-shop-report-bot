import logging

from keywordchecker.generic import handle_invalid_keywords

VALID_TRANSAKSI_KEYWORD = [
    'Dikembalikan',
    'Diterima & Selesai',
    'Diproses Pelapak',
    'Dikirim'
]

VALID_NOMINAL_REMIT_KEYWORD = [
    # Nominal Remit
    'Remit untuk transaksi',
]

VALID_KEUNTUNGAN_TAMBAHAN_KEYWORD = [
    # Keuntungan Tambahan
]

VALID_KERUGIAN_TAMBAHAN_KEYWORD = [
    # Kerugian Tambahan
    'Pemotongan biaya Super Seller',
    'Pembayaran fee promo campaign',
    'Pemotongan biaya kelebihan berat',
]

VALID_BONUS_KEYWORD = [
    # BisaBonus
    'Injection Promo Cashback Sicepat',
]

VALID_SALDO_KEYWORD = [
    # Need to Check
    'Pembayaran pinjaman',
    'Pembayaran topup DANA',
    'Pembayaran untuk tagihan',
    'Pencairan dana',

    'Penambahan dana',
    'Refund transaksi',
] + VALID_NOMINAL_REMIT_KEYWORD + VALID_KEUNTUNGAN_TAMBAHAN_KEYWORD + VALID_KERUGIAN_TAMBAHAN_KEYWORD + VALID_BONUS_KEYWORD

def check_saldo_keyword(bl_file, df):
    logging.debug("Check BisaSaldo Keyword in {0}".format(bl_file))

    invalid_rows = df[~df['Keterangan'].str.contains('|'.join(VALID_SALDO_KEYWORD))]

    handle_invalid_keywords('BisaSaldo', bl_file, invalid_rows)


def check_status_keyword(bl_file, df):
    logging.debug("Check BisaTransaksi Keyword in {0}".format(bl_file))

    invalid_rows = df[~df['Status'].isin(VALID_TRANSAKSI_KEYWORD)]

    handle_invalid_keywords('BisaTransaksi', bl_file, invalid_rows)
