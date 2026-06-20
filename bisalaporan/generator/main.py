# Configure Config
import logging

import process.bukalapak.v2 as bukalapak_v2
import process.shopee.v2 as shopee_v2
import process.shopee.v3 as shopee_v3
import process.tokopedia.v1 as tokopedia_v1
import process.tiktok.v1 as tiktok_v1
import process.tokopedia.v2 as tokopedia_v2
from final.generic import generate_final
from process.preprocess import generate_report_list
from utility.generic import ignore_warning

logging.basicConfig(level=logging.INFO)

# Marketplace -> ordered list of processor modules (Tiktok / Shopee first).
MARKETPLACE_PROCESSORS = {
    'tiktok': [tiktok_v1],
    'shopee': [shopee_v2, shopee_v3],
    'tokopedia': [tokopedia_v1, tokopedia_v2],
    'bukalapak': [bukalapak_v2],
}


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


def main():
    logging.debug("Start Main Process")

    # PreProcess
    ignore_warning(True)
    list_report = generate_report_list(False)

    # Process Marketplace
    run(list_report)


if __name__ == "__main__":
    logging.info("Start Application")
    main()
