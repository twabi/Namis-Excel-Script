import pylightxl as xl
from tkinter import Tk  # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename
import sys
import pandas as pd
import datetime
from fuzzywuzzy import fuzz
import numpy as np
from fuzzywuzzy import process

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
  newDateFormat = str(dt.year) + str(dt.month).zfill(2) + str(dt.day).zfill(2)
  return newDateFormat

def get_ratio(row):
    name = row['reporting unit']
    average = 0
    for orgName in orgUnits:
        partial = fuzz.partial_ratio(name, orgName[0])
        ratio = fuzz.ratio(name,  orgName[0])
        tokenSort = fuzz.token_sort_ratio(name, orgName[0])

        average = (partial + ratio + tokenSort)/3
    return average
def put_rain_id(x):
    id = "qXPtzeZ6Pge"
    return id

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
    #pd.set_option('display.max_rows', None)
    df = pd.read_csv(path)
    df['period'] = df['period'].apply(change_date)
    #excel_data_df['reporting unit'] = excel_data_df['reporting unit'].apply(put_id)

    choices = df['reporting unit'].unique()
    for index, orgName in enumerate(orgUnits):
        df['reporting unit'] = np.where(df['reporting unit'] == orgName[0], orgName[1], df['reporting unit'])
        possibilities = process.extract(str(orgName[0]), choices, limit=100, scorer=fuzz.ratio)
        #df['reporting unit'] = np.where(df['reporting unit'] == orgName[0], orgName[1], df['reporting unit'])
        variate = [possible[0] for possible in possibilities if possible[1] > 73]
        if len(variate) != 0:
            print(variate[0])
            df['reporting unit'] = np.where(df['reporting unit'] == variate[0], orgName[1], df['reporting unit'])

    print(df)

    columns = ['dataelement', 'period', 'orgunit', 'catoptcombo', 'attroptcombo', 'value']
    index = range(0, len(df.index))
    edf = pd.DataFrame(index=index, columns=columns)
    edf['dataelement'] = edf['dataelement'].apply(put_rain_id)
    edf['period'] = df['period'].values
    edf["orgunit"] = df['reporting unit'].values
    edf["value"] = df['value'].values
    print(edf)

    editedFileName = filePath + "{}.csv".format(newFileName)
    print(editedFileName)
    edf.to_csv(editedFileName, index=False)

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

