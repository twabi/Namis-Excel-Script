import pylightxl as xl
from tkinter import Tk  # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename
import sys
import pandas as pd
import datetime
from fuzzywuzzy import fuzz
import numpy as np

Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
filename = askopenfilename()  # show an "Open" dialog box and return the path to the selected file

fileName = filename.split("/")
last = len(fileName) - 1
name = fileName[last].split(".")[0]
newFileName = "New-" + name

filePath = filename.replace(fileName[last], "")

od = xl.readcsv(fn="organisationUnits.csv", delimiter=',')
orgNames = od.ws(ws='Sheet1').range(address='A1:B100', formula=False)
orgNames.pop(0)

orgUnits = []
[orgUnits.append(x) for x in orgNames if x not in orgUnits]


def change_date(x):
  dateString = x
  dt = datetime.datetime.strptime(dateString, '%Y-%m-%d')
  newDateFormat = str(dt.year) + str(dt.month) + str(dt.day)
  return newDateFormat

def put_id(data):
    for orgUnit in orgUnits:
        # if there is a match
        orgString = str(orgUnit[0]).strip('"')
        partial = fuzz.partial_ratio(str(data).lower(), orgString.lower())
        ratio = fuzz.ratio(str(data).lower(), orgString.lower())
        tokenSort = fuzz.token_sort_ratio(str(data).lower(), orgString.lower())


        if data.lower().replace(" ", "") in orgString.lower().replace(" ", ""):
            print(data.lower(), orgString.lower())
            data = orgUnit[1]
        elif ratio > 90 and partial > 70 and tokenSort > 70:
            print(data.lower(), orgString.lower())
            data = orgUnit[1]
    return data

def reStructure(path):
    df = pd.read_csv(path)
    df['period'] = df['period'].apply(change_date)
    #excel_data_df['reporting unit'] = excel_data_df['reporting unit'].apply(put_id)

    string = "Neno Boma"
    id = "bjasbxiabns"

    #partial = fuzz.partial_ratio(df['reporting unit'].item().tolist(), string.lower())
    #ratio = fuzz.ratio(df['reporting unit'].item(), string.lower())
    #tokenSort = fuzz.token_sort_ratio(df['reporting unit'].item(), string.lower())

    df['reporting unit'] = np.where(
        df['reporting unit'] == string
        , id, df['reporting unit'])
    #for orgUnit in orgUnits:
        #df['reporting unit'] = df['reporting unit'].replace([str(orgUnit[0])],str(orgUnit[0]))

    print(df)


if filename.lower().endswith('.csv'):
    reStructure(filename)

elif filename.lower().endswith('.xls'):

    # Read and store content of an excel file
    read_file = pd.read_excel(filename)

    # Write the dataframe object into csv file
    newPath = filePath + "{}.csv".format(name)
    read_file.to_csv(newPath,
                     index=None,
                     header=True)
    reStructure(newPath)

else:
    print("invalid file format", file=sys.stderr)

