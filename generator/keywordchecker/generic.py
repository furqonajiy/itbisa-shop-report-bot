import logging


def handle_invalid_keywords(file_type, file_name, invalid_rows):
    if not invalid_rows.empty:
        raise ValueError("Check {0} Keyword failed in {1}: {2}".format(file_type, file_name, invalid_rows.to_string()))
    else:
        logging.debug("All keywords valid")
