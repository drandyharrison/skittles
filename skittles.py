import pandas as pd
from XLSXhandler import XLSXhandler

# config
# TODO read config from a JSON
xlsx_fname = "Victory Buoys Fixtures 2018-2019.xlsx"
team = 'Victory Buoys'

xlsx = XLSXhandler(xlsx_fname)

# Read the skittles Excel
try:
    # load spreadsheet
    if xlsx.get_xlsx_from_file():
        print(xlsx.get_sheet_names())
        # TODO get Victory Buoys fixtures
except FileNotFoundError:
    print("ERROR: file not found - {}".format(fname))

# TODO write (future) fixtures to Google calendar

