import os

# Repo root = three levels up from generator/utility/constant.py
_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Default input/output folders (repo-relative, OS-independent).
# Inputs (marketplace exports) live under data/; generated reports under reports/.
DEFAULT_DATA_DIR = os.path.join(_REPO_ROOT, 'data')
DEFAULT_REPORTS_DIR = os.path.join(_REPO_ROOT, 'reports')

# Marketplace name (as it appears in the report filename) -> reports/<subfolder>
MARKETPLACE_FOLDERS = {
    'Shopee': 'shopee',
    'Tiktok': 'tiktokshop',
    'Tokopedia': 'tokopedia',
    'Bukalapak': 'bukalapak',
}

_data_dir = DEFAULT_DATA_DIR
_reports_dir = DEFAULT_REPORTS_DIR


def set_dirs(data_dir=None, reports_dir=None):
    """Override the default input/output folders (used by the CLI)."""
    global _data_dir, _reports_dir
    if data_dir:
        _data_dir = os.path.abspath(data_dir)
    if reports_dir:
        _reports_dir = os.path.abspath(reports_dir)


def get_data_dir():
    """Return the folder that input marketplace exports are read from."""
    return _data_dir


def get_reports_dir():
    """Return the folder that generated reports are written under."""
    return _reports_dir
