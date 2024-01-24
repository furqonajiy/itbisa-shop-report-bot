import logging


def generate_bisafee(shp_file, df):
    logging.info("Generate BisaFee Shopee from {0} ({1} rows)".format(shp_file, len(df)))

    df = df[['No. Pesanan', 'Harga Asli Produk', 'Jumlah Pengembalian Dana ke Pembeli', 'Total Diskon Produk', 'Ongkir Dibayar Pembeli', 'Diskon Ongkir Ditanggung Jasa Kirim', 'Gratis Ongkir dari Shopee', 'Ongkir yang Diteruskan oleh Shopee ke Jasa Kirim', 'Biaya Administrasi', 'Biaya Layanan (termasuk PPN 11%)', 'Total Penghasilan']]
    df.columns = ['Invoice', 'Harga Asli Produk', 'Jumlah Pengembalian Dana ke Pembeli', 'Total Diskon Produk', 'Ongkir Dibayar Pembeli', 'Diskon Ongkir Ditanggung Jasa Kirim', 'Gratis Ongkir dari Shopee', 'Ongkir yang Diteruskan oleh Shopee ke Jasa Kirim', 'Biaya Administrasi', 'Biaya Layanan (termasuk PPN 11%)', 'Nominal Remit']

    df['Potongan Pembayaran (Fee)'] = -(df['Ongkir yang Diteruskan oleh Shopee ke Jasa Kirim'])
    df['Keuntungan Tambahan (Fee)'] = (df['Ongkir Dibayar Pembeli'] + df['Diskon Ongkir Ditanggung Jasa Kirim'] + df['Gratis Ongkir dari Shopee'] + df['Ongkir yang Diteruskan oleh Shopee ke Jasa Kirim'])
    df['Kerugian Tambahan (Fee)'] = -(df['Biaya Administrasi'] + df['Biaya Layanan (termasuk PPN 11%)'])
    df = df[['Invoice', 'Potongan Pembayaran (Fee)', 'Nominal Remit', 'Keuntungan Tambahan (Fee)', 'Kerugian Tambahan (Fee)']]

    return df
