import unittest
import skittles


class testSkittles(unittest.TestCase):
    def SetUp(self):
        """set-up code, which is called before each test, to avoid repetition"""
        pass

    def tearDown(self):
        """code to tidy up after each test"""
        pass

    def test_get_fixtures_xlsx_not_string(self):
        """Test that get_fixtures throws a ValueError for a non-string Excel filename"""
        print("@test_get_fixtures_xlsx_not_string")
        # arrange
        fname = 25
        team = "blank"
        # act
        # assert
        self.assertRaises(ValueError, skittles.get_fixtures, fname, team)

    def test_get_fixtures_xlsx_blank(self):
        """check get_fixtures throws a ValueError for a blank string Excel filename"""
        print("@test_get_fixtures_xlsx_blank")
        # arrange
        fname = "   "
        team = "blank"
        # act
        # assert
        self.assertRaises(ValueError, skittles.get_fixtures, fname, team)

    def test_get_fixtures_xlsx_empty(self):
        """check get_fixtures throws a ValueError for an empty string Excel filename"""
        print("@test_get_fixtures_xlsx_empty")
        # arrange
        fname = ""
        team = "blank"
        # act
        # assert
        self.assertRaises(ValueError, skittles.get_fixtures, fname, team)

    def test_get_fixtures_team_not_string(self):
        """Test that get_fixtures throws a ValueError for a non-string team name"""
        print("@test_get_fixtures_team_not_string")
        # arrange
        fname = "blank.xlsx"
        team = 25
        # act
        # assert
        self.assertRaises(ValueError, skittles.get_fixtures, fname, team)

    def test_get_fixtures_team_blank(self):
        """check get_fixtures throws a ValueError for a blank team string"""
        print("@test_get_fixtures_team_blank")
        # arrange
        fname = "blank.xlsx"
        team = "   "
        # act
        # assert
        self.assertRaises(ValueError, skittles.get_fixtures, fname, team)

    def test_get_fixtures_team_empty(self):
        """check get_fixtures throws a ValueError for an empty team string"""
        print("@test_get_fixtures_team_empty")
        # arrange
        fname = "blank.xlsx"
        team = ""
        # act
        # assert
        self.assertRaises(ValueError, skittles.get_fixtures, fname, team)

    def test_get_fixtures_not_xlsx_suffix(self):
        """check get_fixtures throws a ValueError for filename string that doesn't have an *.xlsx suffix"""
        print("@test_get_fixtures_team_empty")
        # arrange
        fname = "blank.xls"
        team = "blank"
        # act
        # assert
        self.assertRaises(ValueError, skittles.get_fixtures, fname, team)


# run tests
if __name__ == '__main__':
    unittest.main(verbosity=2)
