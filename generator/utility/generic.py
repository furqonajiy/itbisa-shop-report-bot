import logging
import os
import warnings

from pandas.core.common import SettingWithCopyWarning


def ignore_warning(ignore):
    logging.debug("Ignore Warning")

    if ignore:
        warnings.simplefilter(action='ignore', category=SettingWithCopyWarning)
        warnings.simplefilter(action='ignore', category=FutureWarning)
        warnings.simplefilter(action='ignore', category=UserWarning)


def create_directory(folder_path):
    if not (os.path.exists(folder_path)):
        logging.debug("{0} not exist, create directory".format(folder_path))
        os.makedirs(folder_path)
