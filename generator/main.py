# Configure Config
import logging

import process.bukalapak.v2 as bukalapak_v2
import process.shopee.v2 as shopee_v2
import process.tokopedia.v1 as tokopedia_v1
import process.tokopedia.v2 as tokopedia_v2
from process.preprocess import generate_report_list
from utility.generic import ignore_warning

logging.basicConfig(level=logging.INFO)


def main():
    logging.debug("Start Main Process")

    # PreProcess
    ignore_warning(True)
    list_report = generate_report_list(False)

    # Process Marketplace
    bukalapak_v2.process(list_report)
    tokopedia_v1.process(list_report)
    tokopedia_v2.process(list_report)
    shopee_v2.process(list_report)


if __name__ == "__main__":
    logging.info("Start Application")
    main()
