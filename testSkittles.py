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


# run tests
if __name__ == '__main__':
    unittest.main(verbosity=2)
