import glob
import logging
import os

from utility.constant import get_data_dir


def generate_report_list(show):
    logging.info("Generate Report List")

    # Collect every Excel/CSV under data/ (recursively, so it works whether the
    # files sit directly in data/ or inside marketplace subfolders). Each
    # processor filters by the marketplace/version token in the filename.
    data_dir = get_data_dir()
    list_report = []
    for pattern in ('*.xls*', '*.csv'):
        list_report += glob.glob(os.path.join(data_dir, '**', pattern), recursive=True)

    list_report = sorted(set(list_report))

    # Check all reports
    if show:
        logging.info("List Report (%d files in %s):", len(list_report), data_dir)
        for report in list_report:
            logging.info(report)

    return list_report
