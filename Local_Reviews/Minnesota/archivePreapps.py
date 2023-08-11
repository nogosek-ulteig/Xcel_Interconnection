import os
import shutil
import sys
import warnings
import glob
import zipfile

warnings.filterwarnings("ignore", category=DeprecationWarning)
warnings.filterwarnings('ignore')

pathToPreapps = r"C:\Users\joe.nogosek\Documents\Projects\NSP_Preapps\Completed"
pathToPreappsArchive = r"C:\Users\joe.nogosek\Documents\Projects\NSP_Preapps\Archive"

count = 0

subs = os.listdir(pathToPreapps)
for sub in subs:
    print(sub)
    feeders = os.listdir(os.path.join(pathToPreapps,sub))
    for feeder in feeders:
        print(" " + feeder)
        cases = os.listdir(os.path.join(pathToPreapps,sub,feeder))
        for case in cases:
            print("  " + case)
            if not os.path.exists(os.path.join(pathToPreappsArchive,sub,feeder,case)):
                shutil.copytree(os.path.join(pathToPreapps,sub,feeder,case),os.path.join(pathToPreappsArchive,sub,feeder,case))
                count += 1

    print("\n" + str(count) + " files moved over!\n")

shutil.rmtree(pathToPreapps)
os.makedirs(pathToPreapps)

print("\nAll finished up! In total " + str(count) + " files were moved over.")
