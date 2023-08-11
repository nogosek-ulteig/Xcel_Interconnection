import sys
import os
from tkinter import *
import subprocess

window=Tk()

window.title("Joe's Script Arsenal")
window.geometry('524x390')

# NSP
def NSP_trackerUpdate():
    os.system(r'python -i  C:\Users\joe.nogosek\Documents\Local_Reviews\Minnesota\trackerAutomation_(without_folders).py')

def NSP_setupFolders():
    os.system(r'python -i  C:\Users\joe.nogosek\Documents\Local_Reviews\Minnesota\setupFolders.py')

def NSP_moveFolders():
    os.system(r'python -i  C:\Users\joe.nogosek\Documents\Local_Reviews\Minnesota\moveFolders.py')

def NSP_batchProcessOneWindow():
    os.system(r'python -i  C:\Users\joe.nogosek\Documents\Local_Reviews\Minnesota\batchProcessOneWindow.py')

def NSP_batchProcess():
    os.system(r'python -i  C:\Users\joe.nogosek\Documents\Local_Reviews\Minnesota\batchProcess.py')

def NSP_TPS():
    os.system(r'python -i  C:\Users\joe.nogosek\Documents\Local_Reviews\Minnesota\autoTPS.py')

def NSP_initialsTrackerUpdate():
    os.system(r'python -i C:\Users\joe.nogosek\Documents\Local_Reviews\Minnesota\initialTrackerUpdate.py')

def NSP_setupInitialReviews():
    os.system(r'python -i  C:\Users\joe.nogosek\Documents\Local_Reviews\Minnesota\setupInitialReviews.py')

def NSP_autoProcessInitials():
    os.system(r'python -i  C:\Users\joe.nogosek\Documents\Local_Reviews\Minnesota\batch_InitialReviews.py')

def NSP_archiveInitials():
    os.system(r'python -i  C:\Users\joe.nogosek\Documents\Local_Reviews\Minnesota\archiveInitials.py')    

def NSP_replaceChecklists():
    os.system(r'python -i  C:\Users\joe.nogosek\Documents\Local_Reviews\Minnesota\replaceChecklists.py')

def NSP_updatePreappTracker():
    os.system(r'python -i  C:\Users\joe.nogosek\Documents\Local_Reviews\Minnesota\updatePreappTracker.py')

def NSP_setupPreapps():
    os.system(r'python -i  C:\Users\joe.nogosek\Documents\Local_Reviews\Minnesota\setupPreapps.py')

def NSP_batchPreapps():
    os.system(r'python -i  C:\Users\joe.nogosek\Documents\Local_Reviews\Minnesota\batch_Preapps.py')

def NSP_archivePreapps():
    os.system(r'python -i  C:\Users\joe.nogosek\Documents\Local_Reviews\Minnesota\archivePreapps.py')


NSP_trackerUpdate_btn = Button(window,text="Update Tracker",bg="green",fg="white",command=NSP_trackerUpdate,width="36")
NSP_trackerUpdate_btn.grid(row=0,column=0)

NSP_setupFolders_btn = Button(window,text="Setup Folders",bg="green",fg="white",command=NSP_setupFolders,width="36")
NSP_setupFolders_btn.grid(row=1,column=0)

NSP_moveFolders_btn = Button(window,text="Move Folders",bg="green",fg="white",command=NSP_moveFolders,width="36")
NSP_moveFolders_btn.grid(row=2,column=0)

NSP_batchProcessOneWindow_btn = Button(window,text="Batch Process (One Window)",bg="green",fg="white",command=NSP_batchProcessOneWindow,width="36")
NSP_batchProcessOneWindow_btn.grid(row=3,column=0)

NSP_batchProcess_btn = Button(window,text="Batch Process (Multiple Windows)",bg="green",fg="white",command=NSP_batchProcess,width="36")
NSP_batchProcess_btn.grid(row=4,column=0)

NSP_TPS_btn = Button(window,text="TPS",bg="green",fg="white",command=NSP_TPS,width="36")
NSP_TPS_btn.grid(row=5,column=0)

NSP_initialsTrackerUpdate_btn = Button(window,text="Update Initials Tracker",bg="green",fg="white",command=NSP_initialsTrackerUpdate,width="36")
NSP_initialsTrackerUpdate_btn.grid(row=6,column=0)

NSP_setupInitialReviews_btn = Button(window,text="Setup Initial Reviews",bg="green",fg="white",command=NSP_setupInitialReviews,width="36")
NSP_setupInitialReviews_btn.grid(row=7,column=0)

NSP_autoProcessInitials_btn = Button(window,text="Batch Process Initial Reviews",bg="green",fg="white",command=NSP_autoProcessInitials,width="36")
NSP_autoProcessInitials_btn.grid(row=8,column=0)

NSP_archiveInitials_btn = Button(window,text="Archive Initial Reviews",bg="green",fg="white",command=NSP_archiveInitials,width="36")
NSP_archiveInitials_btn.grid(row=9,column=0)

NSP_replaceChecklists_btn = Button(window,text="Replace Checklists",bg="green",fg="white",command=NSP_replaceChecklists,width="36")
NSP_replaceChecklists_btn.grid(row=10,column=0)

NSP_updatePreappTracker_btn = Button(window,text="Update Preapp Tracker",bg="green",fg="white",command=NSP_updatePreappTracker,width="36")
NSP_updatePreappTracker_btn.grid(row=11,column=0)

NSP_setupPreapps_btn = Button(window,text="Setup Preapps",bg="green",fg="white",command=NSP_setupPreapps,width="36")
NSP_setupPreapps_btn.grid(row=12,column=0)

NSP_batchPreapps_btn = Button(window,text="Batch Process Preapps",bg="green",fg="white",command=NSP_batchPreapps,width="36")
NSP_batchPreapps_btn.grid(row=13,column=0)

NSP_archivePreapps_btn = Button(window,text="Archive Preapps",bg="green",fg="white",command=NSP_archivePreapps,width="36")
NSP_archivePreapps_btn.grid(row=14,column=0)

#WI
def WI_trackerUpdate():
    os.system(r'python -i  C:\Users\joe.nogosek\Documents\Local_Reviews\Wisconsin\trackerAutomation_WI.py')

def WI_setupFolders():
    os.system(r'python -i  C:\Users\joe.nogosek\Documents\Local_Reviews\Wisconsin\setupFolders_WI.py')

def WI_replaceChecklists():
    os.system(r'python -i  C:\Users\joe.nogosek\Documents\Local_Reviews\Wisconsin\replaceChecklists_WI.py')
    
def WI_moveFolders():
    os.system(r'python -i  C:\Users\joe.nogosek\Documents\Local_Reviews\Wisconsin\moveFolders_WI.py')

def WI_batchProcess():
    os.system(r'python -i  C:\Users\joe.nogosek\Documents\Local_Reviews\Wisconsin\batchProcess_WI_oneWindow.py')
    

WI_trackerUpdate_btn = Button(window,text="Update Tracker",bg="red",fg="white",command=WI_trackerUpdate,width="36")
WI_trackerUpdate_btn.grid(row=6,column=1)

WI_setupFolders_btn = Button(window,text="Setup Folders",bg="red",fg="white",command=WI_setupFolders,width="36")
WI_setupFolders_btn.grid(row=7,column=1)

WI_replaceChecklists_btn = Button(window,text="Replace Checklists",bg="red",fg="white",command=WI_replaceChecklists,width="36")
WI_replaceChecklists_btn.grid(row=8,column=1)

WI_moveFolders_btn = Button(window,text="Move Folders",bg="red",fg="white",command=WI_moveFolders,width="36")
WI_moveFolders_btn.grid(row=9,column=1)

WI_batchProcess_btn = Button(window,text="Batch Process (One Window)",bg="red",fg="white",command=WI_batchProcess,width="36")
WI_batchProcess_btn.grid(row=10,column=1)

# PSCo
def PSCo_updateTracker():
    os.system(r'python -i  C:\Users\joe.nogosek\Documents\Local_Reviews\Colorado\updateTracker_PSCo.py')

def PSCo_setupFolders():
    os.system(r'python -i  C:\Users\joe.nogosek\Documents\Local_Reviews\Colorado\setupFolders_PSCo.py')

def PSCo_batchProcessOneWindow():
    os.system(r'python -i  C:\Users\joe.nogosek\Documents\Local_Reviews\Colorado\batchProcess_PSCo_oneWindow.py')

def PSCo_batchProcess():
    os.system(r'python -i  C:\Users\joe.nogosek\Documents\Local_Reviews\Colorado\batchProcess_PSCo.py')

def PSCo_replaceChecklists():
    os.system(r'python -i  C:\Users\joe.nogosek\Documents\Local_Reviews\Colorado\replaceChecklists_PSCo.py')

def PSCo_moveFolders():
    os.system(r'python -i  C:\Users\joe.nogosek\Documents\Local_Reviews\Colorado\moveFolders.py')

PSCo_updateTracker_btn = Button(window,text="Update Tracker",bg="blue",fg="white",command=PSCo_updateTracker,width="36")
PSCo_updateTracker_btn.grid(row=0,column=1)

PSCo_setupFolders_btn = Button(window,text="Setup Folders",bg="blue",fg="white",command=PSCo_setupFolders,width="36")
PSCo_setupFolders_btn.grid(row=1,column=1)

PSCo_batchProcessOneWindow_btn = Button(window,text="Batch Process (One Window)",bg="blue",fg="white",command=PSCo_batchProcessOneWindow,width="36")
PSCo_batchProcessOneWindow_btn.grid(row=2,column=1)

PSCo_batchProcess_btn = Button(window,text="Batch Process (Multiple Windows)",bg="blue",fg="white",command=PSCo_batchProcess,width="36")
PSCo_batchProcess_btn.grid(row=3,column=1)

PSCo_replaceChecklists_btn = Button(window,text="Replace Checklists",bg="blue",fg="white",command=PSCo_replaceChecklists,width="36")
PSCo_replaceChecklists_btn.grid(row=4,column=1)

PSCo_moveFolders_btn = Button(window,text="Move Folders",bg="blue",fg="white",command=PSCo_moveFolders,width="36")
PSCo_moveFolders_btn.grid(row=5,column=1)

window.mainloop()
