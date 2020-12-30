import pylightxl as xl
from tkinter import Tk     # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename
import sys
import pandas as pd

# we don't want a full GUI, so keep the root window from appearing
Tk().withdraw()
# show an "Open" dialog box and return the path to the selected file
filename = askopenfilename()

#create a new-filename using the old filename from the path obtained
fileName = filename.split("/")
last = len(fileName) - 1
name = fileName[last].split(".")[0]
newFileName = "New-"+ name

#removing the extension from the file path
filePath = filename.replace(fileName[last], "")

def reStructure(path) :
    #read the excel file from dhis2 with market names and their ids
    db = xl.readcsv(fn="DHIS2_Markets.csv", delimiter=',')
    #read the excel file from dhis2 with crop names and their respective ids
    cropsDb = xl.readcsv(fn="dhis2_crops - Sheet1.csv", delimiter=',')

    #get the market name and its respective id list from the dhis2 excel sheet
    marketIDs = db.ws(ws='Sheet1').col(col=1)
    marketNames = db.ws(ws='Sheet1').col(col=2)

    # get the crop name and its respective id list from the dhis2 excel sheet
    crops = cropsDb.ws(ws='Sheet1').col(col=2)
    cropsIDs = cropsDb.ws(ws='Sheet1').col(col=1)

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
    #create an array of indexes based on a key word present in the excel file
    startIndexes = [x for x in range(len(markets)) if markets[x] == "MARKET"]
    endIndexes = [x for x in range(len(markets)) if markets[x] == "AVERAGE PR."]

    #extract the data from the excel file using the starting and ending indexes specified above, as cell ranges
    for x in range(len(endIndexes)) :
        cIndex = startIndexes[x] - 1
        cropAdd = "A" + str(cIndex)
        cropName = md.ws(ws='Sheet1').address(address=cropAdd)

        #replace the crop name from the excel sheet with the id name from dhis2
        for i, crop in enumerate(crops, start=0):
            if cropName.lower() == crop.lower() or (cropName.lower() in crop.lower()) or (
                    crop.lower() in cropName.lower()):
                cropName = cropsIDs[i]

        startIndexes[x] = startIndexes[x] + 2
        startAdd = "A" + str(startIndexes[x])
        endAdd = "E" + str(endIndexes[x])

        addRange = startAdd + ":" + endAdd

        #replace the market names from the excel file selected with the ids from the dhis2 marketIDs list
        crops2 = md.ws(ws='Sheet1').range(address=addRange, formula=False)
        for index, data in enumerate(crops2, start=2):
            for i, marketName in enumerate(marketNames):
                if marketName.lower() == data[0].lower() or (marketName.lower() in data[0].lower()) or (
                        data[0].lower() in marketName.lower()):
                    data[0] = marketIDs[i]

        crops2[0][0] = cropName

        # add it to the global array variable to be used later in the writing the new excel file
        sheetDictionary.append(crops2)

    #removing an overflowing element from the markets list
    for sheet in sheetDictionary:
        sheet.pop(1)

    # take this list for example as our input data that we want to put in column A
    columnHeader = ["Crop", "Period", "Org Unit", "Value"]

    # create a black db
    db = xl.Database()

    #create the array to hold our sheets
    weekArray = sheetDictionary[0][0][1:]

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

    # iterate through every sheet to be created in the document
    for num, week in enumerate(weekArray) :
        #add a blank worksheet to the db for each week
        db.add_ws(ws=week)

        #loop through the crops list to add to the crops column
        for row_id, data in enumerate(cropList, start=2) :
            db.ws(ws=week).update_index(row=row_id, col=1, val=data)

        #loop through the markets list to add to the org units column
        for row_id, data in enumerate(orgUnits, start=2) :
            db.ws(ws=week).update_index(row=row_id, col=3, val=data)

        #for each values in the sheet, add respective values to the values column
        for row_id, data in enumerate(valueList, start=2) :
            db.ws(ws=week).update_index(row=row_id, col=4, val=data[num])

        # write the column headers to the excel sheet
        for col_id, data in enumerate(columnHeader, start=1):
            db.ws(ws=week).update_index(row=1, col=col_id, val=data)

    #write out the document finally
    editedFileName = filePath + "{}.xlsx".format(newFileName)
    print(editedFileName)
    xl.writexl(db=db, fn=editedFileName)

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

