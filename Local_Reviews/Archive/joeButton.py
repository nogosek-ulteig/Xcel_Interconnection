# Made by Joe Nogosek, joe.nogosek@ulteig.com
# 12/4/2021

import openpyxl
from openpyxl.workbook import Workbook
from openpyxl import load_workbook
import warnings
import os
import shutil
import sys

warnings.filterwarnings('ignore')

IA = sys.argv[1]
status = sys.argv[2]
reviewer = sys.argv[3]

if status == 'Approved':
    status = 'Appr'
elif status == 'Rejected':
    status = 'Rej'
else:
    status = 'CA'

pathToFolder = r'G:\2021\21.00016\Reviews\JN\{0}-{1}'.format(IA, reviewer)
newPathToFolder = r'G:\2021\21.00016\Reviews\JN\{0}-{1}-{2}'.format(IA, reviewer, status)
v2Path = r'G:\2021\21.00016\Reviews\JN\{0}-{1}-{2} v2'.format(IA, reviewer, status)
v3Path = r'G:\2021\21.00016\Reviews\JN\{0}-{1}-{2} v3'.format(IA, reviewer, status)
os.rename(pathToFolder, newPathToFolder)

pathToApproved = r'G:\2021\21.00016\Reviews\Complete-Approved'
pathToRejected = r'G:\2021\21.00016\Reviews\Complete-Rejected'
rejectedV2 = os.path.join(pathToRejected, '{0}-{1}-{2}'.format(IA, reviewer, status))
rejectedV3 = os.path.join(pathToRejected, '{0}-{1}-{2} v2'.format(IA, reviewer, status))
if status == 'Appr' or status == 'CA':
    shutil.move(newPathToFolder, pathToApproved)
else:
    if not os.path.exists(rejectedV2) and not os.path.exists(rejectedV3):
        shutil.move(newPathToFolder, pathToRejected)
    elif os.path.exists(rejectedV2) and not os.path.exists(rejectedV3):
        os.rename(newPathToFolder, v2Path)
        newPathToFolder = r'G:\2021\21.00016\Reviews\JN\{0}-{1}-{2} v2'.format(IA, reviewer, status)
        shutil.move(newPathToFolder, pathToRejected)
    else:
        os.rename(newPathToFolder, v3Path)
        newPathToFolder = r'G:\2021\21.00016\Reviews\JN\{0}-{1}-{2} v3'.format(IA, reviewer, status)
        shutil.move(newPathToFolder, pathToRejected)
