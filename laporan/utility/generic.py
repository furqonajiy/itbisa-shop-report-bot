import logging
import os
import warnings

try:  # pandas 1.5-2.x: pandas.errors; older: pandas.core.common; pandas 3.0 removed it (CoW default)
    from pandas.errors import SettingWithCopyWarning
except ImportError:  # pragma: no cover
    try:
        from pandas.core.common import SettingWithCopyWarning
    except ImportError:
        SettingWithCopyWarning = None

from laporan.utility.constant import MARKETPLACE_FOLDERS, get_reports_dir


def ignore_warning(ignore):
    logging.debug("Ignore Warning")

    if ignore:
        if SettingWithCopyWarning is not None:
            warnings.simplefilter(action='ignore', category=SettingWithCopyWarning)
        warnings.simplefilter(action='ignore', category=FutureWarning)
        warnings.simplefilter(action='ignore', category=UserWarning)


def create_directory(folder_path):
    if not (os.path.exists(folder_path)):
        logging.debug("{0} not exist, create directory".format(folder_path))
        os.makedirs(folder_path)


def detect_marketplace(filename):
    """Return the reports/ subfolder for a report filename (e.g. '... Laporan Shopee.xlsx')."""
    for name, folder in MARKETPLACE_FOLDERS.items():
        if name in filename:
            return folder
    raise ValueError("Tidak dapat menentukan marketplace dari file: {0}".format(filename))


def build_report_path(laporan_path):
    """Route a computed Laporan filename into reports/<marketplace>/.

    The per-marketplace generators still compute the report filename themselves
    (Transaksi/Saldo/Fee -> Laporan); this only re-roots that
    filename's directory to reports/<marketplace>/ and creates the folder.
    """
    filename = os.path.basename(laporan_path)
    folder = os.path.join(get_reports_dir(), detect_marketplace(filename))
    create_directory(folder)
    return os.path.join(folder, filename)
