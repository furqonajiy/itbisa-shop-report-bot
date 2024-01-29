import logging

from keywordchecker.generic import handle_invalid_keywords

VALID_SALDO_KEYWORD = [
    'Withdrawal',  # Not Used

    # Need to check
    'Sisa Saldo Mitra dikembalikan',
    'Penggunaan Saldo Tokopedia untuk pembelian',
    'Dipotong karena Solusi dari Resolusi',
    'Dipotong karena Pergantian kurir',
    'Ganti Kurir ke Kurir Non Promo',
    'Penarikan subsidi promo ongkir',

    # Nominal Remit
    'Transaksi Penjualan Berhasil',
    'Pemotongan Ongkir',
    'Pemotongan untuk Asuransi',
    'Pemotongan biaya proteksi produk',
    'Dipotong karena kelebihan ongkos kirim',
    'Selisih ongkos kirim',

    # Keuntungan Tambahan
    'Subsidi Kupon Toko',

    # Kerugian Tambahan
    'Pemotongan Biaya Layanan',
    'Pemotongan Voucher Merchant',
    'Pemotongan Saldo untuk Kupon Toko',
    'Pemotongan Selisih Ongkir',

    # BisaBonus
    'Cashback pengiriman GrabExpress',
    'Cashback atas pengiriman cashless JNE',
    'Cashback atas pengiriman cashless J&T',
    'Cashback atas pengiriman cashless Lion Parcel'
]

VALID_TRANSAKSI_KEYWORD = [
    # V1
    'Transaksi selesai..\nDana akan diteruskan ke penjual.',
    'Transaksi dibatalkan.',
    'Pesanan telah dikirim..\nPesanan dalam proses pengiriman oleh kurir.',
    'Pesanan telah tiba di tujuan..\nDana akan diteruskan ketika barang dikonfirmasi pembeli atau otomatis dalam 48 jam.',
    'Pemesanan sedang diproses oleh penjual.',
    'Menunggu Pick Up',

    # V2
    'Pesanan Diproses',
    'Menunggu Pickup',
    'Pesanan Dikirim',
    'Pesanan Tiba',
    'Pesanan Selesai',
]

VALID_NOMINAL_REMIT_KEYWORD = [
    # Nominal Remit
    'Transaksi Penjualan Berhasil',
    'Pemotongan Ongkir',
    'Pemotongan untuk Asuransi',
    'Pemotongan biaya proteksi produk',
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
    'Selisih ongkos kirim',
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
    'Cashback pengiriman GrabExpress',
    'Cashback atas pengiriman cashless JNE',
    'Cashback atas pengiriman cashless J&T',
    'Cashback atas pengiriman cashless Lion Parcel'
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
