import sys
import os
from tkinter import *
import subprocess

window=Tk()

window.title("Josh's Script GUI")
window.geometry('262x156')

# PSCo
def PSCo_updateTracker():
    os.system(r'python -i  C:\Users\josh.berg\Documents\Local_Reviews\Colorado\updateTracker_PSCo - Copy.py')

def PSCo_setupFolders():
    os.system(r'python -i  C:\Users\josh.berg\Documents\Local_Reviews\Colorado\setupFolders_PSCo - Copy.py')

def PSCo_batchProcessOneWindow():
    os.system(r'python -i  C:\Users\josh.berg\Documents\Local_Reviews\Colorado\batchProcess_PSCo_oneWindow - Copy.py')

def PSCo_batchProcess():
    os.system(r'python -i  C:\Users\josh.berg\Documents\Local_Reviews\Colorado\batchProcess_PSCo - Copy.py')

def PSCo_replaceChecklists():
    os.system(r'python -i  C:\Users\josh.berg\Documents\Local_Reviews\Colorado\replaceChecklists_PSCo - Copy.py')

def PSCo_moveFolders():
    os.system(r'python -i  C:\Users\josh.berg\Documents\Local_Reviews\Colorado\moveFolders - Copy.py')

PSCo_updateTracker_btn = Button(window,text="Update Tracker",bg="blue",fg="white",command=PSCo_updateTracker,width="36")
PSCo_updateTracker_btn.grid(row=0,column=0)

PSCo_setupFolders_btn = Button(window,text="Setup Folders",bg="blue",fg="white",command=PSCo_setupFolders,width="36")
PSCo_setupFolders_btn.grid(row=1,column=0)

PSCo_batchProcessOneWindow_btn = Button(window,text="Batch Process (One Window)",bg="blue",fg="white",command=PSCo_batchProcessOneWindow,width="36")
PSCo_batchProcessOneWindow_btn.grid(row=2,column=0)

PSCo_batchProcess_btn = Button(window,text="Batch Process (Multiple Windows)",bg="blue",fg="white",command=PSCo_batchProcess,width="36")
PSCo_batchProcess_btn.grid(row=3,column=0)

PSCo_replaceChecklists_btn = Button(window,text="Replace Checklists",bg="blue",fg="white",command=PSCo_replaceChecklists,width="36")
PSCo_replaceChecklists_btn.grid(row=4,column=0)

PSCo_moveFolders_btn = Button(window,text="Move Folders",bg="blue",fg="white",command=PSCo_moveFolders,width="36")
PSCo_moveFolders_btn.grid(row=5,column=0)

window.mainloop()
