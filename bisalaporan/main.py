#!/usr/bin/env python3
"""Command-line entry point for the itbisa-bisalaporan report generator.

Reads marketplace exports from data/ and writes Laporan workbooks
(Invoice / Jual / Remit / Bonus sheets) into
reports/<marketplace>/ (reports/shopee, reports/tiktokshop,
reports/tokopedia, reports/bukalapak).

Examples (PowerShell):
    python main.py                       # process every marketplace
    python main.py --shopee --tiktok     # only Shopee and Tiktok
    python main.py --data-dir .\data --output-dir .\reports
    python main.py --show-files -v       # list discovered inputs, debug logs
"""
import argparse
import logging
import os
import sys

# Make the generator/ package importable regardless of the current directory.
GENERATOR_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'generator')
if GENERATOR_DIR not in sys.path:
    sys.path.insert(0, GENERATOR_DIR)

from utility import constant  # noqa: E402  (path setup must run first)
from utility.generic import ignore_warning  # noqa: E402
from process.preprocess import generate_report_list  # noqa: E402
from rekonsiliasi.generic import generate_reconciliation  # noqa: E402
import main as generator  # noqa: E402  (generator/main.py orchestration)

MARKETPLACES = list(generator.MARKETPLACE_PROCESSORS.keys())


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        prog='main.py',
        description='Generate Laporan workbooks from marketplace exports under '
                    'data/ into reports/<marketplace>/.')
    for marketplace in MARKETPLACES:
        parser.add_argument(
            '--{0}'.format(marketplace), action='store_true',
            help='Process {0} only.'.format(marketplace.capitalize()))
    parser.add_argument('--data-dir', default=None,
                        help='Input folder (default: ./data).')
    parser.add_argument('--output-dir', default=None,
                        help='Output folder (default: ./reports).')
    parser.add_argument('--show-files', action='store_true',
                        help='Log the discovered input files.')
    parser.add_argument('--reconcile', action='store_true',
                        help='Write Rekonsiliasi <Marketplace>.xlsx (read-only audit of '
                             'Saldo/Fee vs what is captured); generates no reports.')
    parser.add_argument('--jual-dir', default=None,
                        help='Folder with the itbisa-shop-report-bot *Jual*.xlsx ledger '
                             '(e.g. ..\\itbisa-shop-report-bot\\data); used by --reconcile for '
                             'the Cek Omzet vs Fee sheet. Falls back to re-derived Omzet if unset.')
    parser.add_argument('--cek-jual', action='store_true',
                        help='Reconcile a list of invoices against the Jual ledger to find '
                             'entry bugs; writes reports/shopee/Cek Jual Shopee.xlsx.')
    parser.add_argument('--invoices', default=None,
                        help='Text file (one invoice per line) for --cek-jual; '
                             'defaults to the built-in list.')
    parser.add_argument('-v', '--verbose', action='store_true',
                        help='Enable debug logging.')
    return parser.parse_args(argv)


def selected_marketplaces(args):
    """Return the chosen marketplaces, or None (= all) when no flag is given."""
    chosen = [m for m in MARKETPLACES if getattr(args, m)]
    return chosen or None


def main(argv=None):
    args = parse_args(argv)
    logging.basicConfig(level=logging.INFO)
    logging.getLogger().setLevel(logging.DEBUG if args.verbose else logging.INFO)
    logging.info("Start Application")

    constant.set_dirs(data_dir=args.data_dir, reports_dir=args.output_dir)
    ignore_warning(True)

    chosen = selected_marketplaces(args)

    if args.cek_jual:
        from rekonsiliasi.cekjual import reconcile_invoices, load_invoices
        invoices = load_invoices(args.invoices) if args.invoices else None
        reconcile_invoices(invoices, jual_dir=args.jual_dir)
        logging.info("Selesai cek Jual. Tersimpan di %s", constant.get_reports_dir())
        return

    if args.reconcile:
        recon = None if chosen is None else [m.capitalize() for m in chosen]
        generate_reconciliation(recon, jual_dir=args.jual_dir)
        logging.info("Selesai rekonsiliasi. Tersimpan di %s", constant.get_reports_dir())
        return

    list_report = generate_report_list(args.show_files)
    if not list_report:
        logging.warning("Tidak ada file ditemukan di %s", constant.get_data_dir())
        return

    generator.run(list_report, chosen)
    logging.info("Selesai. Laporan tersimpan di %s", constant.get_reports_dir())


if __name__ == '__main__':
    main()
