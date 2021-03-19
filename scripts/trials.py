import pylightxl as xl
from tkinter import Tk  # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename
import sys
import pandas as pd
import datetime
from fuzzywuzzy import fuzz
import numpy as np
from fuzzywuzzy import process

dateString = "2005-11-02"

dt = datetime.datetime.strptime(dateString, '%Y-%m-%d')
newDateFormat = str(dt.year) + str(dt.month.real).zfill(2) + str(dt.day).zfill(2)

print(newDateFormat)