import pandas as pd
import xlrd

# read the xlsx file use pandas
# TODO why isn't this function appearing as an attribute of the module when it's imported?
# TODO doxygen documentation
# * fname is name of the Excel file to read
def ld_xlsx_pandas(fname) :
    # check filename is not empty
    if not fname:
        print("filename empty")
    else:
        # my first try statmennt
        try:
            # load spreadsheet
            xl_data = pd.ExcelFile(fname)
            # print the sheetnames
            print(xl_data.sheet_names)
        except FileNotFoundError:
            print("ERROR: file not found - {}".format(fname))
