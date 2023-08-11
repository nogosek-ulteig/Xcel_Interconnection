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

pathToFolder = r'G:\2023\23.22984\CO_Reviews\Ready_for_QC\{0}-{1}'.format(IA, reviewer)
newPathToFolder = r'G:\2023\23.22984\CO_Reviews\Ready_for_QC\{0}-{1}-{2}'.format(IA, reviewer, status)
v2Path = r'G:\2023\23.22984\CO_Reviews\Ready_for_QC\{0}-{1}-{2}v2'.format(IA, reviewer, status)
v3Path = r'G:\2023\23.22984\CO_Reviews\Ready_for_QC\{0}-{1}-{2}v3'.format(IA, reviewer, status)
v4Path = r'G:\2023\23.22984\CO_Reviews\Ready_for_QC\{0}-{1}-{2}v4'.format(IA, reviewer, status)
v5Path = r'G:\2023\23.22984\CO_Reviews\Ready_for_QC\{0}-{1}-{2}v5'.format(IA, reviewer, status)
v6Path = r'G:\2023\23.22984\CO_Reviews\Ready_for_QC\{0}-{1}-{2}v6'.format(IA, reviewer, status)
v7Path = r'G:\2023\23.22984\CO_Reviews\Ready_for_QC\{0}-{1}-{2}v7'.format(IA, reviewer, status)
v8Path = r'G:\2023\23.22984\CO_Reviews\Ready_for_QC\{0}-{1}-{2}v8'.format(IA, reviewer, status)
os.rename(pathToFolder, newPathToFolder)

pathToApproved = r'G:\2023\23.22984\CO_Reviews\Complete-Approved'
pathToRejected = r'G:\2023\23.22984\CO_Reviews\Complete-Rejected'
rejectedV2 = os.path.join(pathToRejected, '{0}-{1}-{2}'.format(IA, reviewer, status))
rejectedV3 = os.path.join(pathToRejected, '{0}-{1}-{2}v2'.format(IA, reviewer, status))
rejectedV4 = os.path.join(pathToRejected, '{0}-{1}-{2}v3'.format(IA, reviewer, status))
rejectedV5 = os.path.join(pathToRejected, '{0}-{1}-{2}v4'.format(IA, reviewer, status))
rejectedV6 = os.path.join(pathToRejected, '{0}-{1}-{2}v5'.format(IA, reviewer, status))
rejectedV7 = os.path.join(pathToRejected, '{0}-{1}-{2}v6'.format(IA, reviewer, status))
rejectedV8 = os.path.join(pathToRejected, '{0}-{1}-{2}v7'.format(IA, reviewer, status))
if status == 'Appr' or status == 'CA':
    shutil.move(newPathToFolder, pathToApproved)
else:
    if not os.path.exists(rejectedV2) and not os.path.exists(rejectedV3) and not os.path.exists(rejectedV4) and not os.path.exists(rejectedV5):
        shutil.move(newPathToFolder, pathToRejected)
    elif os.path.exists(rejectedV2) and not os.path.exists(rejectedV3) and not os.path.exists(rejectedV4) and not os.path.exists(rejectedV5):
        os.rename(newPathToFolder, v2Path)
        newPathToFolder = r'G:\2023\23.22984\CO_Reviews\Ready_for_QC\{0}-{1}-{2}v2'.format(IA, reviewer, status)
        shutil.move(newPathToFolder, pathToRejected)
    elif os.path.exists(rejectedV2) and os.path.exists(rejectedV3) and not os.path.exists(rejectedV4) and not os.path.exists(rejectedV5):
        os.rename(newPathToFolder, v3Path)
        newPathToFolder = r'G:\2023\23.22984\CO_Reviews\Ready_for_QC\{0}-{1}-{2}v3'.format(IA, reviewer, status)
        shutil.move(newPathToFolder, pathToRejected)
    elif os.path.exists(rejectedV2) and os.path.exists(rejectedV3) and os.path.exists(rejectedV4) and not os.path.exists(rejectedV5):
        os.rename(newPathToFolder, v4Path)
        newPathToFolder = r'G:\2023\23.22984\CO_Reviews\Ready_for_QC\{0}-{1}-{2}v4'.format(IA, reviewer, status)
        shutil.move(newPathToFolder, pathToRejected)
    elif os.path.exists(rejectedV2) and os.path.exists(rejectedV3) and os.path.exists(rejectedV4) and os.path.exists(rejectedV5) and not os.path.exists(rejectedV6):
        os.rename(newPathToFolder, v5Path)
        newPathToFolder = r'G:\2023\23.22984\CO_Reviews\Ready_for_QC\{0}-{1}-{2}v5'.format(IA, reviewer, status)
        shutil.move(newPathToFolder, pathToRejected)
    elif os.path.exists(rejectedV2) and os.path.exists(rejectedV3) and os.path.exists(rejectedV4) and os.path.exists(rejectedV5) and os.path.exists(rejectedV6) and not os.path.exists(rejectedV7):
        os.rename(newPathToFolder, v6Path)
        newPathToFolder = r'G:\2023\23.22984\CO_Reviews\Ready_for_QC\{0}-{1}-{2}v6'.format(IA, reviewer, status)
        shutil.move(newPathToFolder, pathToRejected)
    elif os.path.exists(rejectedV2) and os.path.exists(rejectedV3) and os.path.exists(rejectedV4) and os.path.exists(rejectedV5) and os.path.exists(rejectedV6) and os.path.exists(rejectedV7) and not os.path.exists(rejectedV8):
        os.rename(newPathToFolder, v7Path)
        newPathToFolder = r'G:\2023\23.22984\CO_Reviews\Ready_for_QC\{0}-{1}-{2}v7'.format(IA, reviewer, status)
        shutil.move(newPathToFolder, pathToRejected)
    else:
        os.rename(newPathToFolder, v8Path)
        newPathToFolder = r'G:\2023\23.22984\CO_Reviews\Ready_for_QC\{0}-{1}-{2}v8'.format(IA, reviewer, status)
        shutil.move(newPathToFolder, pathToRejected)
