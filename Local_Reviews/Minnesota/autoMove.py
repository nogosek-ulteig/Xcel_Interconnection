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

IA = sys.argv[1]
overUnder = sys.argv[2]
reviewer = sys.argv[3]

pathToFolder = r'G:\2023\23.22981\Reviews\{0}\{1}-{2}'.format(reviewer, IA, reviewer)
pathToQC = r'G:\2023\23.22981\Reviews\Ready_for_QC'
shutil.move(pathToFolder, pathToQC)
