import logging

import pandas as pd


def bisainvoice_to_excel(df, path, sheet_name):
    logging.debug("Export BisaInvoice: {0} to Excel".format(path))

    # Convert value to String
    df[['Invoice', 'Ongkir', 'Asuransi']] = df[['Invoice', 'Ongkir', 'Asuransi']].astype(str)  # String

    # Sort based on Invoice
    df = df.sort_values('Invoice')

    # Reset Index
    df = df.reset_index(drop=True)
    df.index = df.index + 1

    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=sheet_name)

        sheet = writer.book[sheet_name]  # Select sheet to be formatted
        sheet.column_dimensions['B'].width = 18  # Tanggal
        sheet.column_dimensions['C'].width = 18  # Marketplace
        sheet.column_dimensions['D'].width = 40  # Invoice
