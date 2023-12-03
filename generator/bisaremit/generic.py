import logging

import pandas as pd
from openpyxl.reader.excel import load_workbook


def bisaremit_to_excel(df, path, sheet_name):
    logging.debug("Export BisaRemit: {0} to Excel".format(path))

    # Reset Index
    df = df.reset_index()
    df.index = df.index + 1

    # Convert Data Type
    df[['Potongan Pembayaran', 'Nominal Remit', 'Keuntungan Tambahan', 'Kerugian Tambahan']] = df[
        ['Potongan Pembayaran', 'Nominal Remit', 'Keuntungan Tambahan', 'Kerugian Tambahan']].astype(str)  # String

    # Export to Excel
    try:
        writer = pd.ExcelWriter(path, mode='a', engine='openpyxl', if_sheet_exists='replace')
        writer.book = load_workbook(path)
        df.to_excel(writer, sheet_name=sheet_name)  # Export
    except FileNotFoundError:
        writer = pd.ExcelWriter(path, engine='openpyxl')
        df.to_excel(writer, sheet_name=sheet_name)  # Export
    except ValueError:
        writer = pd.ExcelWriter(path, engine='openpyxl')
        df.to_excel(writer, sheet_name=sheet_name)  # Export

    sheet = writer.book[sheet_name]  # Select sheet to be formatted
    sheet.column_dimensions['B'].width = 30  # Invoice Bukalapak
    sheet.column_dimensions['C'].width = 18  # Tanggal Remit
    sheet.column_dimensions['D'].width = 22  # Potongan Pembayaran
    sheet.column_dimensions['E'].width = 15  # Nominal Remit
    sheet.column_dimensions['F'].width = 22  # Keuntungan Tambahan
    sheet.column_dimensions['G'].width = 19  # Kerugian Tambahan
    writer.close()
