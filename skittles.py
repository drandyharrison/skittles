import pandas as pd
from XLSXhandler import XLSXhandler

# config
# TODO read config from a JSON
xlsx_fname = "Victory Buoys Fixtures 2018-2019.xlsx"
team = 'Victory Buoys'

xlsx = XLSXhandler(xlsx_fname)

def get_fixtures(xlsx_fname, team):
    """Read the skittles fixtures from Excel
    * xlsx_fname - name of the Excel file containing the fixture information
    * team - the team whose fixtures are to be read (= worksheet name)"""
    # TODO validation
    # TODO xlsx_fname is a non-empty string with the suffix *.xlsx
    if isinstance(xlsx_fname, str):
        # check whether string is empty or blank
        if not (xlsx_fname and xlsx_fname.strip()):
            raise ValueError("@get_fixtures({}, {}) - {} is blank or empty".format(xlsx_fname, team, xlsx_fname))
    else:
        raise ValueError("@get_fixtures({}, {}) - {} is not a string".format(xlsx_fname, team, xlsx_fname))
    # TODO team is a non-empty string
    if not isinstance(team, str):
        raise ValueError("@get_fixtures({}, {}) - {} is not a string".format(xlsx_fname, team, team))
    if xlsx.get_xlsx_from_file():
        print(xlsx.get_sheet_names())
        # TODO get Victory Buoys fixtures

# TODO write (future) fixtures to Google calendar

