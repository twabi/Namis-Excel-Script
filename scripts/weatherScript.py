import pylightxl as xl
from tkinter import Tk     # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename
import sys
import pandas as pd
import datetime
from fuzzywuzzy import fuzz

Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file

fileName = filename.split("/")
last = len(fileName) - 1
name = fileName[last].split(".")[0]
newFileName = "New-"+ name

filePath = filename.replace(fileName[last], "")

def reStructure(path) :
    #read the excel file
    md = xl.readcsv(fn=path, delimiter=',')

    od = xl.readcsv(fn="organisationUnits.csv", delimiter=',')
    orgNames = od.ws(ws='Sheet1').range(address='A1:B100', formula=False)
    orgNames.pop(0)

    orgUnits = []
    [orgUnits.append(x) for x in orgNames if x not in orgUnits]

    #initialize a global variable array to be used through out the document
    sheetDictionary = []

    #get the first column from the file with all the markets
    #markets = md.ws(ws='Sheet1').col(col=1)
    #print(len(markets))
    print(orgUnits)

    whole = md.ws(ws='Sheet1').range(address='A1:C231380', formula=False)
    whole.pop(0)
    print(len(whole))

    for i, data in enumerate(whole):
        dateString = data[1]
        dt = datetime.datetime.strptime(dateString, '%Y-%m-%d')
        newDateFormat = str(dt.year) + str(dt.month)  + str(dt.day)
        data[1] = newDateFormat

    print(whole)
    orgUnits_list_contains(whole, orgUnits)




def orgUnits_list_contains(dataList, orgList):
    # Iterate in the 1st list
    for data in dataList:
        # Iterate in the 2nd list
        for orgUnit in orgList:
            # if there is a match
            orgString = str(orgUnit[0]).strip('"')
            partial = fuzz.partial_ratio(str(data[0]).lower(), orgString.lower())
            ratio = fuzz.ratio(str(data[0]).lower(), orgString.lower())
            tokenSort = fuzz.token_sort_ratio(str(data[0]).lower(), orgString.lower())

            print(data[0])
            print(orgString)

            if data[0].lower().replace(" ", "") in orgString.lower().replace(" ", ""):
                print(data[0].lower(), orgString.lower())
                data[0] = orgUnit[1]
            elif ratio > 90 and partial > 70 and tokenSort > 70:
                print(data[0].lower(), orgString.lower())
                data[0] = orgUnit[1]


if filename.lower().endswith('.csv'):
    reStructure(filename)

elif filename.lower().endswith('.xls') :

    # Read and store content of an excel file
    read_file = pd.read_excel(filename)

    # Write the dataframe object into csv file
    newPath = filePath +  "{}.csv".format(name)
    read_file.to_csv(newPath,
                     index=None,
                     header=True)
    reStructure(newPath)

else:
    print("invalid file format", file=sys.stderr)

