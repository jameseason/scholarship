
import os
import sys
from Tkinter import *
from makeDoc import run
from makeDoc import *
import Tkinter, Tkconstants, tkFileDialog, tkFont

  
### Main Program and GUI ###
def startGUI():
    defaultFolder = ''
    defaultOutput = ''
    defaultAttrRow = 1
    defaultStartRow = 2
        
    w.title("Generate .docx from excel directory")

    # excel file
    excelLabel = Label(w, text="Excel file path (.xlsx):")
    excelLabel.grid(row=0, sticky=W, pady=5, padx=5)

    inputLocation.set(defaultFolder)

    excelEntry = Entry(w, textvariable=inputLocation, width=30)
    excelEntry.grid(row=0,column=1, sticky=W, pady=5, padx=5)

    Button(w, text='Browse', command=setInputDialog).grid(row=0, column=2, sticky=W, pady=5, padx=5)
    
    # attr row
    attrRowLabel = Label(w, text="Row number of attributes in Excel file:")
    attrRowLabel.grid(row=1, sticky=W, pady=5, padx=5)
    
    attrRow.set(defaultAttrRow)
    
    attrRowEntry = Entry(w, textvariable=attrRow, width=5)
    attrRowEntry.grid(row=1, column=1, sticky=W, pady=5, padx=5)
    
    
    # start row 
    startRowLabel = Label(w, text="Row number of first household in Excel file:")
    startRowLabel.grid(row=2, sticky=W, pady=5, padx=5)
    startRow.set(defaultStartRow)
    startRowEntry = Entry(w, textvariable=startRow, width=5)
    startRowEntry.grid(row=2, column=1, sticky=W, pady=5, padx=5)
    
    
    # end row 
    endRowLabel = Label(w, text="Row number of last household in Excel file: \n(leave blank to go to end of file)")
    endRowLabel.grid(row=3, sticky=W, pady=5, padx=5)
    endRow.set('')
    endRowEntry = Entry(w, textvariable=endRow, width=5)
    endRowEntry.grid(row=3, column=1, sticky=W, pady=5, padx=5)
    
    
    
    # output file
    outputLabel = Label(w, text="Output path (.docx):")
    outputLabel.grid(row=4, sticky=W, pady=5, padx=5)
    outputLocation.set(defaultOutput)
    outputEntry = Entry(w, textvariable=outputLocation, width=30)
    outputEntry.grid(row=4,column=1, sticky=W, pady=5, padx=5)

    Button(w, text='Browse', command=setOutputDialog).grid(row=4, column=2, sticky=W, pady=5, padx=5)
    
    Button(w, text='Start', command=run).grid(row=6, column=0, sticky=E, pady=5, padx=5)
    Button(w, text='Quit', command=quit).grid(row=6, column=1, sticky=W, pady=5, padx=5)
    
    titleCheck.set(1)
    Checkbutton(w, text="Include title pages for districts", variable=titleCheck).grid(row=5, column=0, sticky=W, pady=5, padx=5, columnspan=2)
    
    w.mainloop()
    

# Set input file
def setInputDialog(): 
    w.filename = tkFileDialog.askopenfilename(initialdir = os.getcwd(),title = "Select input (xlsx) file",filetypes = (("Excel file","*.xlsx"),("all files","*.*")))
    if w.filename != None and len(w.filename) > 0:
        if w.filename[-5:] != '.xlsx':
            w.filename += '.xlsx'
        inputLocation.set(w.filename)
        
# Set output file():
def setOutputDialog(): 
    w.filename = tkFileDialog.asksaveasfilename(initialdir = os.getcwd(),title = "Select output (docx) file",filetypes = (("Microsoft Word file","*.docx"),("all files","*.*")))
    if w.filename != None and len(w.filename) > 0:
        if w.filename[-5:] != '.docx':
            w.filename += '.docx'
        outputLocation.set(w.filename)
    

 
startGUI()



