# Configure Config
import logging

import process.bukalapak as bukalapak
import process.tokopedia as tokopedia
from process.preprocess import generate_report_list
from utility.generic import ignore_warning

logging.basicConfig(level=logging.INFO)


def main():
    logging.debug("Start Main Process")

    # PreProcess
    ignore_warning(True)
    list_report = generate_report_list(False)

    # Process Marketplace
    bukalapak.process(list_report)
    tokopedia.process(list_report)


if __name__ == "__main__":
    logging.info("Start Application")
    main()
