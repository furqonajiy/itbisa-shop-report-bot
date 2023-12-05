import logging

from keywordchecker.generic import handle_invalid_keywords

VALID_SALDO_KEYWORD = [
    'Withdrawal',  # Not Used

    'Sisa Saldo Mitra dikembalikan',  # Need to check
    'Penggunaan Saldo Tokopedia untuk pembelian',  # Need to check
    'Dipotong karena Solusi dari Resolusi',  # Need to check
    'Dipotong karena Pergantian kurir',  # Need to check
    'Ganti Kurir ke Kurir Non Promo',  # Need to check
    'Penarikan subsidi promo ongkir',  # Need to check

    'Pemotongan Ongkir',  # Nominal Remit
    'Pemotongan untuk Asuransi',  # Nominal Remit
    'Transaksi Penjualan Berhasil',  # Nominal Remit

    'Pemotongan Biaya Layanan',  # Kerugian Tambahan
    'Pemotongan Voucher Merchant',  # Kerugian Tambahan

    'Cashback pengiriman GrabExpress',  # BisaBonus
    'Cashback atas pengiriman cashless JNE',  # BisaBonus
    'Cashback atas pengiriman cashless J&T',  # BisaBonus
    'Cashback atas pengiriman cashless Lion Parcel'  # BisaBonus
]

VALID_TRANSAKSI_KEYWORD = [
    'Transaksi selesai..\nDana akan diteruskan ke penjual.',  # v1

    'Pesanan Diproses',
    'Menunggu Pickup',
    'Pesanan Dikirim',
    'Pesanan Tiba',
    'Pesanan Selesai',
]

VALID_NOMINAL_REMIT_KEYWORD = [
    'Pemotongan Ongkir',  # Nominal Remit
    'Pemotongan untuk Asuransi',  # Nominal Remit
    'Transaksi Penjualan Berhasil',  # Nominal Remit
]

VALID_BONUS_KEYWORD = [
    'Cashback pengiriman GrabExpress',  # BisaBonus
    'Cashback atas pengiriman cashless JNE',  # BisaBonus
    'Cashback atas pengiriman cashless J&T',  # BisaBonus
    'Cashback atas pengiriman cashless Lion Parcel'  # BisaBonus
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
