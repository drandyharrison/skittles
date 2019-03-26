from XLSXhandler import XLSXhandler
from JSONhandler import JSONhandler
from GoogleCalAPIHandler import GoogleCalAPIHandler
import datetime


# if we define argument type, not-defined when pass argument of different type - so still need to do validation
# prefix arguments with a_ to distinguish them from globals with same name
def get_fixtures(a_xlsx_fname:str, a_team:str):
    """Read the skittles fixtures from Excel
    * xlsx_fname - name of the Excel file containing the fixture information
    * team - the team whose fixtures are to be read (= worksheet name)"""
    if isinstance(a_xlsx_fname, str):
        # check whether string is empty or blank
        if not (a_xlsx_fname and a_xlsx_fname.strip()):
            raise ValueError("@get_fixtures({}, {}) - {} is blank or empty".format(a_xlsx_fname, a_team, a_xlsx_fname))
        elif not a_xlsx_fname.endswith(".xlsx"):
            raise ValueError("@get_fixtures({}, {}) - {} has wrong suffix".format(a_xlsx_fname, a_team, a_xlsx_fname))
    else:
        raise ValueError("@get_fixtures({}, {}) - {} is not a string".format(a_xlsx_fname, a_team, a_xlsx_fname))
    if isinstance(a_team, str):
        # check whether string is empty or blank
        if not (a_team and a_team.strip()):
            raise ValueError("@get_fixtures({}, {}) - {} is blank or empty".format(a_xlsx_fname, a_team, a_team))
    else:
        raise ValueError("@get_fixtures({}, {}) - {} is not a string".format(a_xlsx_fname, team, team))
    # get Excel file handler
    xlsx = XLSXhandler(a_xlsx_fname)
    if xlsx.get_xlsx_from_file():
        teams_list = xlsx.get_sheet_names()
        if a_team not in teams_list:
            raise ValueError("@get_fixtures({}, {}) - team not in list {}".format(a_xlsx_fname, a_team, teams_list))
        # get fixtures data
        raw_data = xlsx.xlsx_data.parse(a_team)
        start_row = 0
        end_row = 41
        fixtures = raw_data.values[start_row:end_row, 0:5]
        for lp1 in range(start_row, end_row):
            # convert date field to date objects
            if isinstance(fixtures[lp1][0], datetime.datetime):
                fixtures[lp1][0] = fixtures[lp1][0].date()
            if isinstance(fixtures[lp1][0], str):
                # assumes the format of the date, if it's a string
                # parse date format in string and handle the different ones
                fixtures[lp1][0] = datetime.datetime.strptime(fixtures[lp1][0], '%d/%m/%Y').date()
        return fixtures

# --------- #
# Main body #
# --------- #

# read config from a JSON
jsonhndlr = JSONhandler("skittles_config.json")
if jsonhndlr.read_json():
    # read key values from config file
    xlsx_fname = jsonhndlr.get_val('xlsx_fname')
    team = jsonhndlr.get_val('team')
    # to handle where team name (for worksheet tab) is spelt different to home team in fixtures
    team2 = jsonhndlr.get_val('home_team')

# create Google calendar API handler
calhndlr = GoogleCalAPIHandler()
#calhndlr.get_next_n_appts(10)
# read the fixtures
fixtures = get_fixtures(xlsx_fname, team)
for row in fixtures:
    # discard games in the past
    date_of_game = row[0]
    if date_of_game > datetime.date.today():
        home_team = row[1]
        home_game = (home_team == team2)
        away_team = row[2]
        #competition = row[3]
        venue = row[4]
        # TODO write to calendar
        print("[{}] {} vs {} @ {}".format(date_of_game, home_team, away_team, venue))

# test adding an event
event = {
  'summary': 'Skittles test',
  'location': 'Venue',
  'description': 'Home/Away (Opposition)',
  'start': {
    'dateTime': '2019-03-28T20:00:00',
    'timeZone': 'Europe/London',
  },
  'end': {
    'dateTime': '2019-03-28T23:00:00',
    'timeZone': 'Europe/London',
  },
}

# TODO check the timezones are valid

calhndlr.add_event("iam.andyharrison@gmail.com", event)

