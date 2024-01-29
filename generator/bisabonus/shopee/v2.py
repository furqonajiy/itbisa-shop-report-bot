import logging

import pandas as pd
from openpyxl.reader.excel import load_workbook

from keywordchecker.shopee import VALID_BONUS_KEYWORD


def generate_bisabonus(tkp_file, df):
    logging.info("Generate BisaBonus from {0} ({1} rows)".format(tkp_file, len(df)))

    # Select valid rows based on Deskripsi
    df = df[df['Deskripsi'].str.contains('|'.join(VALID_BONUS_KEYWORD))]
    if len(df) == 0:
        logging.debug("No valid BisaBonus keyword")
        return

    # Initialize Jurnal
    df['Akun Debit'] = 'Kas ITBisa'
    df['Akun Kredit'] = 'Bonus Shopee'

    # Select Needed Column
    df = df[['Date', 'Akun Debit', 'Akun Kredit', 'Nominal (Rp)', 'Deskripsi']]

    # Change Column Name
    df.columns = ['Tanggal', 'Akun Debit', 'Akun Kredit', 'Nominal', 'Keterangan']

    # Export to Existing WorkBook
    path = (tkp_file
            .replace(' v1', '')
            .replace(' v2', '')
            .replace('BisaSaldo', 'BisaLaporan'))

    # Check if file exist
    with pd.ExcelWriter(path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        writer.book = load_workbook(path)
        df.to_excel(writer, sheet_name='BisaBonus Shopee')

        sheet = writer.book['BisaBonus Shopee']  # Select sheet to be formatted
        sheet.column_dimensions['B'].width = 18  # Tanggal
        sheet.column_dimensions['C'].width = 16  # Akun Debit
        sheet.column_dimensions['D'].width = 16  # Akun Kredit
        sheet.column_dimensions['E'].width = 15  # Nominal
