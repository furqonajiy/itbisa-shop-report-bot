import glob
import logging

from utility.constant import BISASALDO_DIR, BISATRANSAKSI_DIR


def generate_report_list(show):
    logging.info("Generate Report List")

    # Generate all report list
    list_report = []
    all_dir = BISASALDO_DIR + BISATRANSAKSI_DIR
    for list_dir in all_dir:
        list_report = list_report + glob.glob(list_dir + '*\*.xls*')
        list_report = list_report + glob.glob(list_dir + '*\*.csv')

    list_report = sorted(list_report)

    # Check all reports
    if show:
        logging.debug("List Report:")
        for report in list_report:
            logging.info(report)

    return list_report
