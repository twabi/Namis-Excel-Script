import pylightxl as xl
from tkinter import Tk     # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename
import sys
import pandas as pd
import datetime
import datetime as gt
import re
import string
import requests
from requests.auth import HTTPBasicAuth
# api-endpoint
url = "https://covmw.com/namistest/api/dataSets/tBglAG7ASSg.json?fields=dataSetElements[dataElement[id,name,formName]]"

alphabet = string.ascii_uppercase
def weeks(year, month):
    _, start, _ = gt.datetime(year, month, 1).isocalendar()
    for d in (31, 30, 29, 28):
        try:
            _, end, _ = gt.datetime(year, month, d).isocalendar()
            break
        except ValueError:  # skip attempts with bad dates
            continue
    if start > 50:  # spillover from previous year
        return list(range(1, end + 1))
    else:
        return list(range(start, end + 1))


# we don't want a full GUI, so keep the root window from appearing
Tk().withdraw()
# show an "Open" dialog box and return the path to the selected file
filename = askopenfilename()

#create a new-filename using the old filename from the path obtained
fileName = filename.split("/")
last = len(fileName) - 1
name = fileName[last].split(".")[0]
newFileName = "New-"+ name

#initialize the year variable for the excel sheet document
#dateYear = ""
#monthNumber = ""

months_short = []
for i in range(1,13):
    months_short.append((i, datetime.date(2008, i, 1).strftime('%b')))
months_choices = []
for i in range(1,13):
    months_choices.append((i, datetime.date(2008, i, 1).strftime('%B')))

year_number = re.findall(r'\d+', name)

month_number = ""
for data in months_choices:
    if data[1].lower() in name.lower():
        print(data[1])
        month_number = data[0]

if month_number == "":
    for data in months_short:
        if data[1].lower() in name.lower():
            print(data[1])
            month_number = data[0]

#removing the extension from the file path
filePath = filename.replace(fileName[last], "")

print(month_number, year_number[0])
weekNumberArray = weeks(int(year_number[0]), int(month_number))
if len(weekNumberArray) == 0:
    tempArray = []
    array2 = weeks(int(year_number[0]), int(month_number)-1)
    for x in range(52 - array2[len(array2) -1]):
        tempArray.append(array2[len(array2) -1] + (x+1))
    weekNumberArray = tempArray

prevArray=[]

for d in weekNumberArray:
    if d in prevArray:
        weekNumberArray.remove(d)
weeksArray = []

#retriving the week numbers for the month name on the document name by cross-checking in the array from the week excel sheet

for x in range(len(weekNumberArray)):
    #modify the elements in the array by adding the year to each element, so as to achieve the dhis2 format
    weeksArray.append(year_number[0] + "W" + str(weekNumberArray[x]))
print(weeksArray)

def reStructure(path) :
    weird_document = False
    start_letter = "A"
    #read the excel file from dhis2 with market names and their ids
    db = xl.readcsv(fn="DHIS2_Markets.csv", delimiter=',')
    #read the excel file from dhis2 with crop names and their respective ids
    # extracting data in json format
    res = requests.get(url, verify=False, auth=HTTPBasicAuth('ahmed', 'Atwabi@20'))
    #print(res)
    apiData = res.json()
    elementArray = apiData["dataSetElements"]
    formNameArray = []
    idArray = []
    for element in elementArray:
        formNameArray.append(element['dataElement']['formName'])
        idArray.append(element['dataElement']['id'])

    #get the market name and its respective id list from the dhis2 excel sheet
    marketIDs = db.ws(ws='Sheet1').col(col=1)
    marketNames = db.ws(ws='Sheet1').col(col=2)

    # get the crop name and its respective id list from the dhis2 excel sheet
    crops = formNameArray
    cropsIDs = idArray

    #remove the column headers from the crop and orgUnits lists
    marketNames.pop(0)
    marketIDs.pop(0)
    crops.pop(0)
    cropsIDs.pop(0)

    #remove the suffix of the market names in the list from "soso market" to just "soso"
    for x in range(len(marketNames)):
        marketNames[x] = marketNames[x].replace(" Market", "")

    #read the excel file
    md = xl.readcsv(fn=path, delimiter=',')

    #initialize a global variable array to be used through out the document
    sheetDictionary = []

    #get the first column from the file with all the markets
    markets = md.ws(ws='Sheet1').col(col=1)
    #print(markets)
    if "ADD" in markets:
        markets = md.ws(ws='Sheet1').col(col=3)
        weird_document = True
    print(markets)
    print(weird_document)
    if weird_document:
        start_letter = "C"

    #create an array of indexes based on a key word present in the excel file
    startIndexes = [x for x in range(len(markets)) if markets[x] == "MARKET"]
    endIndexes = [x for x in range(len(markets)) if markets[x] == "AVERAGE PR."]

    if (len(startIndexes) == 0) and (len(endIndexes) == 0) :
        startIndexes = [x for x in range(len(markets)) if (markets[x] == "" or markets[x] == " ")]
        endIndexes = [x for x in range(len(markets)) if "AVERAGE PRICE" in str(markets[x])]
        for x in startIndexes:
            if x+1 in startIndexes:
                startIndexes.remove(x+1)
            if x+2 in startIndexes:
                startIndexes.remove(x+2)
            if (startIndexes[len(startIndexes) -1] - startIndexes[len(startIndexes) - 2] == 1) or (startIndexes[len(startIndexes) -1] - startIndexes[len(startIndexes) - 2] == 2) or (startIndexes[len(startIndexes) -1] - startIndexes[len(startIndexes) - 2] == 3) :
                startIndexes.pop()
        startIndexes.pop()

    if "Unnamed: 2" in markets:
        startIndexes[0] = startIndexes[0]+2

    cutoff = len(endIndexes)
    print(cutoff)
    overboard = len(startIndexes)
    startIndexes = startIndexes[0:cutoff]
    print(startIndexes)
    print(endIndexes)
    #extract the data from the excel file using the starting and ending indexes specified above, as cell ranges
    for x in range(len(startIndexes)) :
        if "unnamed".lower() in markets[0].lower():
            cIndex = startIndexes[x] + 2
        elif "this year".lower() in markets[0].lower():
            cIndex = startIndexes[x] + 3
        else:
            cIndex = startIndexes[x] - 1

        if markets[0] == "Unnamed: 2" :
            if cIndex == 5:
                cIndex = cIndex + 1
        else:
            if cIndex == 6:
                cIndex = cIndex -1
        cropAdd = start_letter + str(cIndex)
        cropName = md.ws(ws='Sheet1').address(address=cropAdd)
        if cropName == "" or cropName == " ":
            cropAdd = start_letter + str(cIndex+1)
            cropName = md.ws(ws='Sheet1').address(address=cropAdd)
            if cropName == "" or cropName == " ":
                cropAdd = start_letter + str(cIndex + 2)
                cropName = md.ws(ws='Sheet1').address(address=cropAdd)

        print(cropName, cIndex)

        #replace the crop name from the excel sheet with the id name from dhis2
        for i, crop in enumerate(crops, start=0):
            if str(cropName).lower() == str(crop).lower() or (str(cropName).lower() in str(crop).lower()) or (
                    str(crop).lower() in str(cropName).lower()):
                cropName = cropsIDs[i]

        end_letter = alphabet[len(weeksArray)]
        if weird_document:
            end_letter = alphabet[len(weeksArray) + 2]

        startIndexes[x] = startIndexes[x] + 2
        startAdd = start_letter + str(startIndexes[x])
        endAdd = end_letter + str(endIndexes[x])

        addRange = startAdd + ":" + endAdd
        #print(addRange)

        #replace the market names from the excel file selected with the ids from the dhis2 marketIDs list
        crops2 = md.ws(ws='Sheet1').range(address=addRange, formula=False)
        #print("empty?", crops2)
        for index, datum in enumerate(crops2, start=2):
            for i, marketName in enumerate(marketNames):
                if str(marketName).lower() == str(datum[0]).lower() or (str(marketName).lower() in str(datum[0]).lower()) or (
                        str(datum[0]).lower() in str(marketName).lower()):
                    datum[0] = marketIDs[i]

        crops2[0][0] = cropName

        # add it to the global array variable to be used later in the writing the new excel file
        sheetDictionary.append(crops2)

    #removing an overflowing element from the markets list
    for sheet in sheetDictionary:
        sheet.pop(1)

    # take this list for example as our input data that we want to put in column A
    columnHeader = ["Livestock", "Period", "Org Unit", "Value"]

    # create a black db
    db = xl.Database()

    #separating individual column lists from the extracted bulky list
    orgUnits = []
    cropList = []
    valueList = []
    for x in range(len(sheetDictionary)):
        for y in sheetDictionary[x][1:] :
            valueList.append(y[1:])
            orgUnits.append(y[0])

            #making sure the crops List and the org Units are synchronized with the blank cells within them
            if y[0] == '':
                cropList.append("")
            else:
                cropList.append(sheetDictionary[x][0][0])

    print(sheetDictionary)
    print(cropList)
    print(valueList)
    print(orgUnits)

    # create the array to hold our sheets
    weekArray = []
    for x in range(len(valueList[0])) :
        weekArray.append("WK" + str(x+1))
    print(weekArray)
    print(weeksArray)

    # iterate through every sheet to be created in the document
    for num, week in enumerate(weekArray) :
        #add a blank worksheet to the db for each week
        db.add_ws(ws=week)

        #loop through the crops list to add to the crops column
        for row_id, datum in enumerate(cropList, start=2) :
            db.ws(ws=week).update_index(row=row_id, col=1, val=datum)

        # loop through the crops list to add to the crops column
        for row_id, wk in enumerate(cropList, start=2):
            db.ws(ws=week).update_index(row=row_id, col=2, val=weeksArray[num])

        #loop through the markets list to add to the org units column
        for row_id, datum in enumerate(orgUnits, start=2) :
            db.ws(ws=week).update_index(row=row_id, col=3, val=datum)

        #for each values in the sheet, add respective values to the values column
        for row_id, datum in enumerate(valueList, start=2) :
            db.ws(ws=week).update_index(row=row_id, col=4, val=datum[num])

        # write the column headers to the excel sheet
        for col_id, datum in enumerate(columnHeader, start=1):
            db.ws(ws=week).update_index(row=1, col=col_id, val=datum)

    #write out the document finally
    editedFileName = filePath + "{}.xlsx".format(newFileName)
    print(editedFileName)
    xl.writexl(db=db, fn=editedFileName)

def delete_excess_rows(path):
    db = pd.read_csv(path)
    document_length = len(db)
    print(document_length)
    thresh = 2000
    if document_length > thresh:
        print("Document too long! removing excess rows...")
        dp = pd.read_csv(path, skipfooter=document_length - thresh, engine='python')
        dp.to_csv(path, index=False)
        reStructure(path)
    else:
        reStructure(path)

if filename.lower().endswith('.csv'):
    delete_excess_rows(filename)

elif filename.lower().endswith('.xls') :
    # Read and store content of an excel file
    read_file = pd.read_excel(filename)

    # Write the dataframe object into csv file
    newPath = filePath + "{}.csv".format(name)
    read_file.to_csv(newPath,
                     index=None,
                     header=True)
    delete_excess_rows(newPath)

else:
    print("invalid file format", file=sys.stderr)

