import pandas
from XLSXhandler import XLSXhandler

# config (deprecated)
# xlsx_fname = "Victory Buoys Fixtures 2018-2019.xlsx"
# team = 'Victory Buoys'

# read config from a JSON
# TODO create a generic JSON file reader
jsondf = pandas.read_json("skittles_config.json")
if jsondf.values.size < 2:
    raise IndexError("JSON file doesn't have enough parameters")
try:
    xlsx_fname = jsondf['xlsx_fname'][0]
except KeyError as e:
    print("Key error: xlsx_fname")
    exit(0)
try:
    team = jsondf['team'][0]
except KeyError as e:
    print("Key error: team")
    exit(0)

# get Excel file handler
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

