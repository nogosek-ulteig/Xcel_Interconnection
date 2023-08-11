# Made by Joe Nogosek, joe.nogosek@ulteig.com
# 12/3/2021

import openpyxl
from openpyxl.workbook import Workbook
from openpyxl import load_workbook
import warnings
import os
import shutil
import sys

warnings.filterwarnings('ignore')

pathToFile = r'C:\Users\joe.nogosek\Documents\Projects\NSP_Initial_Reviews\JN\Initial Review_Case#0{}.xlsm'.format(str(sys.argv[1]))
pathToQC = r'C:\Users\joe.nogosek\Documents\Projects\NSP_Initial_Reviews\Ready_to_Process\Initial Review_Case#0{}.xlsm'.format(str(sys.argv[1]))

shutil.move(pathToFile, pathToQC)
