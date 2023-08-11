import os
import shutil
import sys
import warnings
import glob
import zipfile

warnings.filterwarnings("ignore", category=DeprecationWarning)
warnings.filterwarnings('ignore')

pathToInitials = r"C:\Users\joe.nogosek\Documents\Projects\NSP_Initial_Reviews\Completed"
pathToInitialsArchive = r"C:\Users\joe.nogosek\Documents\Projects\NSP_Initial_Reviews\Archive"

count = 0

subs = os.listdir(pathToInitials)
for sub in subs:
    print(sub)
    feeders = os.listdir(os.path.join(pathToInitials,sub))
    for feeder in feeders:
        print(" " + feeder)
        cases = os.listdir(os.path.join(pathToInitials,sub,feeder))
        for case in cases:
            print("  " + case)
            if not os.path.exists(os.path.join(pathToInitialsArchive,sub,feeder,case)):
                shutil.copytree(os.path.join(pathToInitials,sub,feeder,case),os.path.join(pathToInitialsArchive,sub,feeder,case))
                count += 1

    print("\n" + str(count) + " files moved over!\n")

shutil.rmtree(pathToInitials)
os.makedirs(pathToInitials)

print("\nAll finished up! In total " + str(count) + " files were moved over.")
