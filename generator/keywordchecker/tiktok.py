import logging

from keywordchecker.generic import handle_invalid_keywords

VALID_TRANSAKSI_KEYWORD = [
    # V1
    'To ship',
    'Shipped',
    'Completed',

    # V2
]

VALID_NOMINAL_REMIT_KEYWORD = [
    # Nominal Remit
    'Transaksi Penjualan Berhasil',
    'Pemotongan Ongkir',
    'Pemotongan untuk Asuransi',
    'Pemotongan biaya proteksi produk',
    'Dipotong karena kelebihan ongkos kirim',
    'Selisih ongkos kirim',
]

VALID_POTONGAN_PEMBAYARAN_KEYWORD = [
    # Nominal Remit
    'Pemotongan Ongkir',
    'Pemotongan untuk Asuransi',
    'Pemotongan biaya proteksi produk',
]

VALID_KEUNTUNGAN_TAMBAHAN_KEYWORD = [
    # Keuntungan Tambahan
    'Subsidi Kupon Toko',
    'Pengembalian dana atas ongkir pengganti',
    'Subsidi Promo Toko',
]

VALID_KERUGIAN_TAMBAHAN_KEYWORD = [
    # Kerugian Tambahan
    'Pemotongan Biaya Layanan',
    'Pemotongan Voucher Merchant',
    'Pemotongan Saldo untuk Kupon Toko',
    'Pemotongan Saldo untuk Promo Toko',
    'Pemotongan Selisih Ongkir',
    'Dipotong karena Solusi dari Resolusi',
    'Dipotong karena Pergantian kurir',
    'Ganti Kurir ke Kurir Non Promo',
]

VALID_BONUS_KEYWORD = [
    # BisaBonus
    'Cashback pengiriman GrabExpress',
    'Cashback atas pengiriman cashless JNE',
    'Cashback atas pengiriman cashless J&T',
    'Cashback atas pengiriman cashless Lion Parcel',

    # BisaBonus - Negatif
    'Penggunaan Saldo Tokopedia untuk pembelian',
    'Penarikan subsidi promo ongkir penyelesaian kendala',
]

VALID_SALDO_KEYWORD = [
    # Not Used
    'Withdrawal',
    'Penarikan Otomatis',
    'Sisa Saldo Mitra dikembalikan',
] + VALID_NOMINAL_REMIT_KEYWORD + VALID_POTONGAN_PEMBAYARAN_KEYWORD + VALID_KEUNTUNGAN_TAMBAHAN_KEYWORD + VALID_KERUGIAN_TAMBAHAN_KEYWORD + VALID_BONUS_KEYWORD

def check_saldo_keyword(tkp_file, df):
    logging.debug("Check BisaSaldo Keyword in {0}".format(tkp_file))

    invalid_rows = df[~df['Description'].str.contains('|'.join(VALID_SALDO_KEYWORD))]

    handle_invalid_keywords('BisaSaldo', tkp_file, invalid_rows)


def check_status_keyword(version, tkp_file, df):
    logging.debug("Check BisaTransaksi Keyword in {0}".format(tkp_file))

    if version == "1":
        invalid_rows = df[~df['Order Status'].isin(VALID_TRANSAKSI_KEYWORD)]

    handle_invalid_keywords('BisaTransaksi', tkp_file, invalid_rows)
