import glob
import logging
import os

import pandas as pd
from openpyxl.reader.excel import load_workbook
from openpyxl.styles import Alignment

from utility.constant import MARKETPLACE_FOLDERS, get_reports_dir

FINAL_SHEET = 'Final'

# Money columns rendered as integers (blank cells stay blank, not 0/NaN).
_MONEY_COLUMNS = [
    'Ongkir', 'Asuransi', 'Omzet\nBarang', 'Nominal\nInvoice',
    'Potongan Pembayaran', 'Nominal Remit', 'Keuntungan\nTambahan', 'Kerugian Tambahan',
]

# Remit amount columns summed when an order has several remit entries.
_REMIT_AMOUNT_COLUMNS = ['Potongan Pembayaran', 'Nominal Remit', 'Keuntungan Tambahan', 'Kerugian Tambahan']

# (header, width) for each Final column, in order.
_COLUMN_LAYOUT = [
    ('Tanggal\nPesan', 20), ('Marketplace', 14), ('Invoice', 40), ('Ongkir', 12),
    ('Asuransi', 12), ('Omzet\nBarang', 15), ('Nominal\nInvoice', 16), ('Tanggal\nRemit', 18),
    ('Potongan Pembayaran', 20), ('Nominal Remit', 15), ('Keuntungan\nTambahan', 14),
    ('Kerugian Tambahan', 18), ('Cek\nRemit', 12), ('Untung Lainnya', 16),
    ('Rugi Lainnya', 14), ('Keterangan', 30),
]


def _to_number(series):
    """Coerce an Excel-read column (often stored as text) to numeric."""
    return pd.to_numeric(series, errors='coerce')


def _read(xls, sheet_name):
    """Return a sheet from an open ExcelFile, or None if it is absent."""
    if sheet_name in xls.sheet_names:
        return xls.parse(sheet_name)
    return None


def _list_workbooks(folder):
    files = glob.glob(os.path.join(folder, '*.xlsx'))
    # Skip Excel temp lock files (~$...).
    return sorted(f for f in files if not os.path.basename(f).startswith('~'))


def _combined_remit(workbooks, remit_sheet):
    """Union every workbook's BisaRemit sheet, one aggregated row per Invoice.

    Lets an order be matched to a remit that landed in a different period's
    BisaLaporan (ordered this month, remitted next month).
    """
    frames = []
    for path in workbooks:
        with pd.ExcelFile(path) as xls:
            remit = _read(xls, remit_sheet)
        if remit is None or remit.empty or 'Invoice' not in remit.columns:
            continue
        keep = ['Invoice', 'Tanggal Remit'] + _REMIT_AMOUNT_COLUMNS
        remit = remit[[c for c in keep if c in remit.columns]].copy()
        remit['Invoice'] = remit['Invoice'].astype(str)
        for col in _REMIT_AMOUNT_COLUMNS:
            if col in remit.columns:
                remit[col] = _to_number(remit[col])
        frames.append(remit)

    if not frames:
        return None

    allremit = pd.concat(frames, ignore_index=True)
    agg = {col: 'sum' for col in _REMIT_AMOUNT_COLUMNS if col in allremit.columns}
    if 'Tanggal Remit' in allremit.columns:
        agg['Tanggal Remit'] = 'max'  # latest remit date (ISO strings sort correctly)
    return allremit.groupby('Invoice', as_index=False).agg(agg)


def _build_final(invoice_df, jual_df, remit_idx):
    """Join one workbook's BisaInvoice with its BisaJual omzet and the remit index."""
    df = invoice_df.copy()
    df['Invoice'] = df['Invoice'].astype(str)
    for col in ['Ongkir', 'Asuransi']:
        df[col] = _to_number(df[col]).fillna(0)

    # Omzet Barang = sum of BisaJual Omzet per Invoice.
    if jual_df is not None and {'Invoice', 'Omzet'}.issubset(jual_df.columns):
        omzet = jual_df[['Invoice', 'Omzet']].copy()
        omzet['Invoice'] = omzet['Invoice'].astype(str)
        omzet['Omzet'] = _to_number(omzet['Omzet']).fillna(0)
        omzet = omzet.groupby('Invoice', as_index=False)['Omzet'].sum()
        df = df.merge(omzet, on='Invoice', how='left')
    else:
        df['Omzet'] = 0
    df['Omzet'] = df['Omzet'].fillna(0)

    # Nominal Invoice = Omzet Barang + Ongkir + Asuransi.
    df['Nominal Invoice'] = df['Omzet'] + df['Ongkir'] + df['Asuransi']

    # Remit columns (left join; blank when no remit is found in any period).
    if remit_idx is not None:
        df = df.merge(remit_idx, on='Invoice', how='left')
    for col in ['Tanggal Remit'] + _REMIT_AMOUNT_COLUMNS:
        if col not in df.columns:
            df[col] = pd.NA

    # Assemble in the requested column order; Cek Remit + the "Lainnya"/Keterangan
    # columns are left blank for manual entry.
    out = pd.DataFrame({
        'Tanggal\nPesan': df['Tanggal'].values,
        'Marketplace': df['Marketplace'].values,
        'Invoice': df['Invoice'].values,
        'Ongkir': df['Ongkir'].values,
        'Asuransi': df['Asuransi'].values,
        'Omzet\nBarang': df['Omzet'].values,
        'Nominal\nInvoice': df['Nominal Invoice'].values,
        'Tanggal\nRemit': df['Tanggal Remit'].values,
        'Potongan Pembayaran': df['Potongan Pembayaran'].values,
        'Nominal Remit': df['Nominal Remit'].values,
        'Keuntungan\nTambahan': df['Keuntungan Tambahan'].values,
        'Kerugian Tambahan': df['Kerugian Tambahan'].values,
        'Cek\nRemit': pd.NA,
        'Untung Lainnya': pd.NA,
        'Rugi Lainnya': pd.NA,
        'Keterangan': pd.NA,
    })

    # Render money as integers; unmatched remit cells stay blank (Int64 <NA>).
    for col in _MONEY_COLUMNS:
        out[col] = out[col].round().astype('Int64')

    return out.sort_values('Invoice').reset_index(drop=True)


def _write_final(out, path, sheet_name=FINAL_SHEET):
    """Append (or replace) the Final sheet in an existing BisaLaporan workbook."""
    try:
        writer = pd.ExcelWriter(path, mode='a', engine='openpyxl', if_sheet_exists='replace')
        writer.book = load_workbook(path)
    except (FileNotFoundError, ValueError):
        writer = pd.ExcelWriter(path, engine='openpyxl')

    out.to_excel(writer, sheet_name=sheet_name, index=False)

    sheet = writer.book[sheet_name]
    for offset, (_, width) in enumerate(_COLUMN_LAYOUT):
        sheet.column_dimensions[chr(ord('A') + offset)].width = width
    for cell in sheet[1]:  # wrap the two-line header labels
        cell.alignment = Alignment(wrap_text=True, vertical='center')
    sheet.freeze_panes = 'A2'

    writer.close()


def generate_final(marketplace):
    """Write the Final sheet into every BisaLaporan workbook of one marketplace.

    marketplace: the report token used in the sheet/file names and as a
    MARKETPLACE_FOLDERS key ('Shopee' / 'Tiktok' / 'Tokopedia' / 'Bukalapak').
    """
    subfolder = MARKETPLACE_FOLDERS.get(marketplace)
    if subfolder is None:
        return
    folder = os.path.join(get_reports_dir(), subfolder)
    if not os.path.isdir(folder):
        return
    workbooks = _list_workbooks(folder)
    if not workbooks:
        return

    invoice_sheet = 'BisaInvoice {0}'.format(marketplace)
    jual_sheet = 'BisaJual {0}'.format(marketplace)
    remit_sheet = 'BisaRemit {0}'.format(marketplace)

    # Build the cross-period remit index once, before writing any Final sheet.
    remit_idx = _combined_remit(workbooks, remit_sheet)

    for path in workbooks:
        with pd.ExcelFile(path) as xls:
            invoice_df = _read(xls, invoice_sheet)
            jual_df = _read(xls, jual_sheet)
        if invoice_df is None or invoice_df.empty or 'Invoice' not in invoice_df.columns:
            continue  # no orders to anchor a Final sheet on
        out = _build_final(invoice_df, jual_df, remit_idx)
        _write_final(out, path)
        logging.info("Generate Final {0} -> {1} ({2} rows)".format(
            marketplace, os.path.basename(path), len(out)))
