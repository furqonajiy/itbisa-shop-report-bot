import logging


def generate_bisafee(shp_file, df):
    logging.info("Generate BisaFee Shopee from {0} ({1} rows)".format(shp_file, len(df)))

    df = df[['No. Pesanan', 'Harga Asli Produk', 'Total Diskon Produk', 'Biaya Administrasi', 'Biaya Layanan (termasuk PPN 11%)', 'Total Penghasilan']]
    df.columns = ['Invoice', 'Harga Asli Produk', 'Total Diskon Produk', 'Biaya Administrasi', 'Biaya Layanan (termasuk PPN 11%)', 'Nominal Remit']

    df['Biaya Admin'] = df['Total Diskon Produk'] + df['Biaya Administrasi'] + df['Biaya Layanan (termasuk PPN 11%)']
    df = df[['Invoice', 'Nominal Remit', 'Biaya Admin']]

    return df
