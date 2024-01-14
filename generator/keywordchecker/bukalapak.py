import logging

from keywordchecker.generic import handle_invalid_keywords

VALID_SALDO_KEYWORD = [
    'Pembayaran pinjaman',  # Not Used
    'Pembayaran topup DANA',  # Not Used
    'Pembayaran untuk tagihan',  # Not Used
    'Pencairan dana',  # Not Used

    'Penambahan dana',  # Need to Check
    'Refund transaksi',  # Need to Check

    'Remit untuk transaksi',  # Nominal Remit

    'Pemotongan biaya Super Seller',  # Kerugian Tambahan
    'Pembayaran fee promo campaign',  # Kerugian Tambahan

    'Injection Promo Cashback Sicepat',  # BisaBonus
    'Pemotongan biaya kelebihan berat'  # Kerugian Tambahan
]

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


def check_saldo_keyword(bl_file, df):
    logging.debug("Check BisaSaldo Keyword in {0}".format(bl_file))

    invalid_rows = df[~df['Keterangan'].str.contains('|'.join(VALID_SALDO_KEYWORD))]

    handle_invalid_keywords('BisaSaldo', bl_file, invalid_rows)


def check_status_keyword(bl_file, df):
    logging.debug("Check BisaTransaksi Keyword in {0}".format(bl_file))

    invalid_rows = df[~df['Status'].isin(VALID_TRANSAKSI_KEYWORD)]

    handle_invalid_keywords('BisaTransaksi', bl_file, invalid_rows)
