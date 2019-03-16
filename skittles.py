import pandas
from XLSXhandler import XLSXhandler
from JSONhandler import JSONhandler

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
        # get fixtures data
        raw_data = xlsx.xlsx_data.parse(team)
        start_row = 0
        end_row = 41
        fixtures = raw_data.values[start_row:end_row, 0:5]
        for lp1 in range(start_row, end_row):
            # TODO check the type of the dates (first column)
            print("[{}] {}\t{}\t{}\t{}".format(fixtures[lp1][0], fixtures[lp1][1], fixtures[lp1][2], fixtures[lp1][3],
                                               fixtures[lp1][4]))
        return fixtures

# TODO write (future) fixtures to Google calendar


# ---------
# Main body
# ---------

# read config from a JSON
jsonhndlr = JSONhandler("skittles_config.json")
if jsonhndlr.read_json():
    # read key values from config file
    xlsx_fname = jsonhndlr.get_val('xlsx_fname')
    team = jsonhndlr.get_val('team')

# get Excel file handler
xlsx = XLSXhandler(xlsx_fname)
# read the fixtures
fixtures = get_fixtures(xlsx_fname, team)

