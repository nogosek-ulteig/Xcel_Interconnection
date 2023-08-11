# Made by Joe Nogosek; joe.nogosek@ulteig.com
# 12/10/2022

import openpyxl
from openpyxl import load_workbook
import os, os.path
from os.path import join
import glob

list_of_files = glob.glob(r"C:\Users\joe.nogosek\Desktop\adf\*") # * means all if need specific format then *.csv

for file in list_of_files:
    new_file = file[:-2] + "RK"
    os.rename(file,new_file)
