# Namis-Excel-Script
## Operation
This is a script to restructure market data for the NAMIS platform by checking the market name and data into their IDs for easy upload onto the DHIS2 platform.

*TO TEST, USE THE SAMPLE FILE IN THE SCRIPTS FOLDER*

Basically, the script to be run is in the scripts folder, called "excelScript.py"
It can be run by simply typing "python excelScript.py" in terminal.

A file picker/selector window pops up by which a user can then just pick the preferred market excel file to be restructured.

Note : The file chosen should be in xls or csv format. The script currently does not support any other formats (e.g. xlsx or xlsm). Also note that the file
needs to be of the markets structure for Namis, any other excel files will obviously not work

After the file has been picked, the script will run and a new excel file(.xlsx) will be created in the directory of the old excel file. If the file chosen is of xls format, the script will convert it to csv for easy reading and such.

Any questions or additions etc, contact itwabi@gmail.com

cheers!
