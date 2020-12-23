import pylightxl as xl
import os

#print(os.getcwd());
#db = xl.readxl(fn='/home/tobirama/Documents/wk APRI 1999.csv')
md = xl.readcsv(fn='/home/tobirama/Documents/wk APRI 1999.csv', delimiter=',')
names = md.ws_names

crop = md.ws(ws='Sheet1').address(address='A5')
markets = md.ws(ws='Sheet1').col(col=1)

print(markets)
