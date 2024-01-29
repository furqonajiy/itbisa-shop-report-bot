import logging

import pandas as pd
from openpyxl.reader.excel import load_workbook


def bisajual_to_excel(df, path, sheet_name):
    logging.debug("Export BisaJual: {0} to Excel".format(path))

    # Convert Data Type
    df = df.astype(str)  # String

    # Sort based on Invoice
    df = df.sort_values(['Invoice', 'SKU'])

    # Reset Index
    df = df.reset_index(drop=True)
    df.index = df.index + 1

    # Export to Excel
    with pd.ExcelWriter(path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        writer.book = load_workbook(path)
        df.to_excel(writer, sheet_name=sheet_name)  # Export

        sheet = writer.book[sheet_name]  # Select sheet to be formatted
        sheet.column_dimensions['B'].width = 50  # SKU
        sheet.column_dimensions['C'].width = 40  # Invoice
        sheet.column_dimensions['D'].width = 10  # Banyak
        sheet.column_dimensions['E'].width = 18  # Omzet
