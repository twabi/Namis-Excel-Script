import pylightxl as xl
import os

#print(os.getcwd());
md = xl.readcsv(fn='/home/tobirama/Documents/wk APRI 1999.csv', delimiter=',')
names = md.ws_names

#crop = md.ws(ws='Sheet1').address(address='A5')
sheetDictionary = []
markets = md.ws(ws='Sheet1').col(col=1)

startIndexes = [x for x in range(len(markets)) if markets[x] == "MARKET"]
endIndexes = [x for x in range(len(markets)) if markets[x] == "AVERAGE PR."]

for x in range(len(startIndexes)) :
    cIndex = startIndexes[x] - 1
    cropAdd = "A" + str(cIndex)
    cropName = md.ws(ws='Sheet1').address(address=cropAdd)

    startIndexes[x] = startIndexes[x] + 2
    startAdd = "A" + str(startIndexes[x])
    endAdd = "C" + str(endIndexes[x])
    addRange = startAdd+ ":" + endAdd

    crops = md.ws(ws='Sheet1').range(address=addRange, formula=False)
    crops[0][0] = cropName
    print(crops)
    sheetDictionary.append(crops)


print(sheetDictionary)


# take this list for example as our input data that we want to put in column A
#mydata = [10,20,30,40]

# create a black db
#db = xl.Database()

# add a blank worksheet to the db
#db.add_ws(ws="Sheet1")

# loop to add our data to the worksheet
#for row_id, data in enumerate(mydata, start=1):
    #db.ws(ws="Sheet1").update_index(row=row_id, col=1, val=data)

# write out the db
#xl.writexl(db=db, fn="/home/tobirama/Documents/output.xlsx")
