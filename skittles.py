import XLSXhandler
import JSONhandler
import GoogleCalAPIHandler
import datetime
import math

def are_get_fixtures_params_valid(a_xlsx_fname:str, a_team:str):
    """Check params for get_fxitures are all valid
    :param xlsx_fname: str
        name of the Excel file containing the fixture information
    :param team: str
        the team whose fixtures are to be read (= worksheet name)
    :raises ValueError: when a parameter has an invalid value
    :return: bool
        are parameters valid
    """
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
    return True

# if we define argument type, not-defined when pass argument of different type - so still need to do validation
# prefix arguments with a_ to distinguish them from globals with same name
def get_fixtures(a_xlsx_fname:str, a_team:str):
    """Read the skittles fixtures from Excel
    :param xlsx_fname: str
        name of the Excel file containing the fixture information
    :param team: str
        the team whose fixtures are to be read (= worksheet name)
    :raises ValueError: when a parameter has an invalid value
    :return: ndarray
        list of fixtures
    """
    if are_get_fixtures_params_valid(a_xlsx_fname, a_team):
        # get Excel file handler
        xlsx = XLSXhandler.XLSXhandler(a_xlsx_fname)
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
    else:
        return None

# --------- #
# Main body #
# --------- #

# read config from a JSON
jsonhndlr = JSONhandler.JSONhandler("skittles_config.json")
if jsonhndlr.read_json():
    # read key values from config file
    xlsx_fname = jsonhndlr.get_val('xlsx_fname')
    team = jsonhndlr.get_val('team')
    # to handle where team name (for worksheet tab) is spelt different to home team in fixtures
    team2 = jsonhndlr.get_val('home_team')
    calendar = jsonhndlr.get_val('calendar')

# create Google calendar API handler
calhndlr = GoogleCalAPIHandler.GoogleCalAPIHandler()
# event template
event = {
    'summary': 'Skittles',
    'location': 'Venue',
    'description': '',
    'start': {
        'dateTime': '2019-03-28T20:00:00',
        'timeZone': 'Europe/London',
    },
    'end': {
        'dateTime': '2019-03-28T23:00:00',
        'timeZone': 'Europe/London',
    },
}
# read the fixtures
fixtures = get_fixtures(xlsx_fname, team)
for row in fixtures:
    # discard games in the past
    date_of_game = row[0]
    if date_of_game > datetime.date.today():
        home_team = row[1]
        home_game = (home_team == team2)
        away_team = row[2]
        competition = row[3]
        venue = row[4]
        # Add game to calendar?
        if isinstance(venue, str):
            add = not(venue in ["BYE"])
        elif isinstance(venue, float):
            add = not math.isnan(venue)
        if add:
            # create event based on fixture data
            suffix = (away_team if home_game else home_team)
            event['summary'] = "Skittles (" + suffix + ")"
            event["location"] = venue
            event["start"]["dateTime"] = str(date_of_game) + "T20:30:00"
            event["end"]["dateTime"] = str(date_of_game) + "T23:30:00"
            # write to calendar
            calhndlr.add_event(calendar, event)
            print("[{}] {} vs {} @ {}".format(date_of_game, home_team, away_team, venue))

