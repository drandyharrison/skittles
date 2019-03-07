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
    if isinstance(xlsx_fname, str):
        # check whether string is empty or blank
        if not (xlsx_fname and xlsx_fname.strip()):
            raise ValueError("@get_fixtures({}, {}) - {} is blank or empty".format(xlsx_fname, team, xlsx_fname))
        elif not xlsx_fname.endswith(".xlsx"):
            raise ValueError("@get_fixtures({}, {}) - {} has wrong suffix".format(xlsx_fname, team, xlsx_fname))
    else:
        raise ValueError("@get_fixtures({}, {}) - {} is not a string".format(xlsx_fname, team, xlsx_fname))
    if isinstance(team, str):
        # check whether string is empty or blank
        if not (team and team.strip()):
            raise ValueError("@get_fixtures({}, {}) - {} is blank or empty".format(xlsx_fname, team, team))
    else:
        raise ValueError("@get_fixtures({}, {}) - {} is not a string".format(xlsx_fname, team, team))
    if xlsx.get_xlsx_from_file():
        teams_list = xlsx.get_sheet_names()
        print(teams_list)
        if team not in teams_list:
            raise ValueError("@get_fixtures({}, {}) - team not in list {}".format(xlsx_fname, team, teams_list))
        # TODO get Victory Buoys fixtures

# TODO write (future) fixtures to Google calendar

# ---------
# Main body
# ---------
get_fixtures(xlsx_fname, team)

