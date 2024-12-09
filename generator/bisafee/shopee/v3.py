import logging


def generate_bisafee(shp_file, df_fee, df_adjust):
    logging.info("Generate BisaFee Shopee from {0} ({1} rows)".format(shp_file, len(df_fee)))

    df_fee = df_fee[['No. Pesanan', 'Harga Asli Produk', 'Jumlah Pengembalian Dana ke Pembeli', 'Total Diskon Produk', 'Ongkir Dibayar Pembeli', 'Diskon Ongkir Ditanggung Jasa Kirim',
                     'Gratis Ongkir dari Shopee', 'Ongkir yang Diteruskan oleh Shopee ke Jasa Kirim', 'Ongkos Kirim Pengembalian Barang',
                     'Biaya Administrasi', 'Biaya Layanan (termasuk PPN 11%)', 'Total Penghasilan']]
    df_fee.columns = ['Invoice', 'Harga Asli Produk', 'Jumlah Pengembalian Dana ke Pembeli', 'Total Diskon Produk', 'Ongkir Dibayar Pembeli', 'Diskon Ongkir Ditanggung Jasa Kirim',
                      'Gratis Ongkir dari Shopee', 'Ongkir yang Diteruskan oleh Shopee ke Jasa Kirim', 'Ongkos Kirim Pengembalian Barang',
                      'Biaya Administrasi', 'Biaya Layanan (termasuk PPN 11%)', 'Nominal Remit']

    df_fee['Potongan Pembayaran (Fee)'] = -(df_fee['Ongkir yang Diteruskan oleh Shopee ke Jasa Kirim'])
    df_fee['Keuntungan Tambahan (Fee)'] = (df_fee['Ongkir Dibayar Pembeli'] + df_fee['Diskon Ongkir Ditanggung Jasa Kirim'] + df_fee['Gratis Ongkir dari Shopee'] + df_fee['Ongkir yang Diteruskan oleh Shopee ke Jasa Kirim'] + + df_fee['Ongkos Kirim Pengembalian Barang'])
    df_fee['Kerugian Tambahan (Fee)'] = -(df_fee['Jumlah Pengembalian Dana ke Pembeli'] + df_fee['Biaya Administrasi'] + df_fee['Biaya Layanan (termasuk PPN 11%)'])
    df_fee = df_fee[['Invoice', 'Potongan Pembayaran (Fee)', 'Nominal Remit', 'Keuntungan Tambahan (Fee)', 'Kerugian Tambahan (Fee)']]

    if len(df_adjust) > 0:
        df_adjust = df_adjust[['No. Pesanan Terhubung', 'Biaya Penyesuaian']].fillna(0)
        df_adjust.columns = ['Invoice', 'Nominal Remit (Fee)']

        df_fee = df_fee.merge(df_adjust, on=['Invoice'], how='left').fillna(0)
        df_fee['Nominal Remit (Fee)'] = df_fee['Nominal Remit (Fee)'].astype(int)

    return df_fee
