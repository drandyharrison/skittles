import ash_utils
import datetime
import math
import pandas

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
        raise ValueError("@get_fixtures({}, {}) - {} is not a string".format(a_xlsx_fname, a_team, a_team))
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
        xlsx = ash_utils.XLSXhandler(a_xlsx_fname)
        if xlsx.get_xlsx_from_file():
            teams_list = xlsx.get_sheet_names()
            if a_team not in teams_list:
                raise ValueError("@get_fixtures({}, {}) - team not in list {}".format(a_xlsx_fname, a_team, teams_list))
            # get fixtures data
            hdr_row = 2
            start_row = 0
            end_row = 40
            xlsx.extract_worksheet_data(a_team, hdr_row, -1, 1+hdr_row+start_row, 1+hdr_row+end_row, 5)
            raw_fixtures = xlsx.get_extracted_as_pandas()
            # create the date index
            date_index = []
            for lp1 in range(start_row, end_row):
                # convert date field to date objects
                if isinstance(raw_fixtures.iloc[lp1][0], datetime.datetime):
                    date_index.append(raw_fixtures.iloc[lp1][0].date())
                if isinstance(raw_fixtures.iloc[lp1][0], str):
                    date_index.append(datetime.datetime.strptime(raw_fixtures.iloc[lp1][0], '%d/%m/%Y').date())
            # create fixtures with date index and oopy in fixture data
            fixtures = pandas.DataFrame(index=date_index, columns=raw_fixtures.columns[1:len(raw_fixtures.columns)], dtype=str)
            for i in range(0, len(raw_fixtures.index)):
                for j in range(1, len(raw_fixtures.columns)):
                    fixtures.iloc[i, j-1] = raw_fixtures.iloc[i, j]
            return fixtures
    else:
        return None

# ------------- #
# Main function #
# ------------- #

def main():
    # read config from a JSON
    jsonhndlr = ash_utils.JSONhandler("skittles_config.json")
    # TODO catch errors when using handlers
    if jsonhndlr.read_json():
        # read key values from config file
        xlsx_fname = jsonhndlr.get_val('xlsx_fname')
        team = jsonhndlr.get_val('team')
        # to handle where team name (for worksheet tab) is spelt different to home team in fixtures
        team2 = jsonhndlr.get_val('home_team')
        calendar = jsonhndlr.get_val('calendar')

    # create Google calendar API handler
    calhndlr = ash_utils.GoogleCalAPIHandler()
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
    for index, row in fixtures.iterrows():
        # discard games in the past
        date_of_game = index
        if date_of_game > datetime.date.today():
            home_team = row['Home Team']
            home_game = (home_team == team2)
            away_team = row['Away Team']
            competition = row['Competition']
            venue = row['Venue']
            # Add game to calendar?
            if isinstance(venue, str):
                add = not(venue in ["BYE"]) and not(competition in ['no match', 'tba'])
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
                if calhndlr.add_event(calendar, event):
                    print("[{}] {} vs {} @ {}".format(date_of_game, home_team, away_team, venue))

if __name__ == "__main__":
    main()