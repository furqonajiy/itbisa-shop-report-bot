#!/usr/bin/env python3
"""Command-line entry point + orchestration for the Laporan report generator.

Reads marketplace exports from data/ and writes Laporan workbooks
(Invoice / Jual / Remit / Bonus sheets) into reports/<marketplace>/
(reports/shopee, reports/tiktokshop, reports/tokopedia, reports/bukalapak).

This is the `laporan` package's orchestration module. Invoke it from the repo
root as a module (`python -m laporan`), or in-process via the bot's
`python main.py --laporan` (which calls main() directly).

Examples (PowerShell, from the repo root):
    python -m laporan                       # process every marketplace
    python -m laporan --shopee --tiktok     # only Shopee and Tiktok
    python -m laporan --data-dir .\\laporan\\data --output-dir .\\laporan\\reports
    python -m laporan --show-files -v       # list discovered inputs, debug logs
"""
import argparse
import logging

import laporan.process.bukalapak.v2 as bukalapak_v2
import laporan.process.shopee.v2 as shopee_v2
import laporan.process.shopee.v3 as shopee_v3
import laporan.process.tokopedia.v1 as tokopedia_v1
import laporan.process.tiktok.v1 as tiktok_v1
import laporan.process.tokopedia.v2 as tokopedia_v2
from laporan.final.generic import generate_final
from laporan.process.preprocess import generate_report_list
from laporan.rekonsiliasi.generic import generate_reconciliation
from laporan.utility import constant
from laporan.utility.generic import ignore_warning

# Marketplace -> ordered list of processor modules (Tiktok / Shopee first).
MARKETPLACE_PROCESSORS = {
    'tiktok': [tiktok_v1],
    'shopee': [shopee_v2, shopee_v3],
    'tokopedia': [tokopedia_v1, tokopedia_v2],
    'bukalapak': [bukalapak_v2],
}

MARKETPLACES = list(MARKETPLACE_PROCESSORS.keys())


def run(list_report, marketplaces=None):
    """Run the selected marketplace processors over the given report list.

    marketplaces: iterable of keys from MARKETPLACE_PROCESSORS, or None for all.
    """
    if marketplaces is None:
        marketplaces = list(MARKETPLACE_PROCESSORS.keys())

    for marketplace in marketplaces:
        for processor in MARKETPLACE_PROCESSORS[marketplace]:
            processor.process(list_report)
        # All of this marketplace's workbooks now exist; build the Final sheet
        # (joins Invoice + Jual + a cross-period Remit lookup).
        generate_final(marketplace.capitalize())


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
                             '(e.g. ..\\data); used by --reconcile for '
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
        from laporan.rekonsiliasi.cekjual import reconcile_invoices, load_invoices
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

    run(list_report, chosen)
    logging.info("Selesai. Laporan tersimpan di %s", constant.get_reports_dir())


if __name__ == '__main__':
    main()
