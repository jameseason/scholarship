# James Eason
# Generate excel file: filename | abstract title | abstract content

import os
import sys
from openpyxl import Workbook
from Tkinter import *

# Remove any non utf-8 characters that might confuse openpyxl
def cleanse(txt):
    return txt.decode('utf-8', 'ignore').encode("utf-8")  

# Search all subfolders in path and add paths of all files to list
def getFiles(path, files):
    if not os.path.isdir(path):
        raise RuntimeError("Path " + path + "is not a directory")     
    dir = os.listdir(path)
    for file in dir:
        newpath = path + "/" + file
        if os.path.isdir(newpath):
            files = getFiles(newpath, files)
        else:
            if '.txt' in newpath:
                files.append(newpath)
    return files
    
# Get topics from topics file
def getTopics(topicsPath):
    topics = {}
    rawContents = cleanse(open(topicsPath, 'r').read()).split('\n')
    for line in rawContents:
        line = line.strip()
        if len(line) < 1 or line[0] == '#':
            pass
        else:
            topicName = line[0:line.index('(')].strip()
            items = line[line.index('(')+1:line.index(')')].split(',')
            for x in range(len(items)):
                items[x] = items[x].strip()
            topics[topicName] = items
    return topics
            
# Get abstracts from a list of files
# Returns dict with years as key and array containing abstracts
def getAbstracts(files):
    abstracts = {}
    for x in range(len(files)):
        fileName = cleanse(files[x])[files[x].rfind('/')+1:-4]
        fileYear = fileName.strip()[:4]
        fileContents = cleanse(open(files[x], 'r').read())
        if not fileYear in abstracts:
            abstracts[fileYear] = []
        abstracts[fileYear].append(fileContents)
    return abstracts
        
# Get how many times a topic appears in a specified year
def getTopicCount(topicItems, abstracts):
    count = 0
    for abstract in abstracts:
        abstract = abstract.lower().split()
        present = False
        for item in topicItems:
            if isPresent(item.lower(), abstract):
                present = True
        if present:
            count += 1
    return count
    
# Check if a topic item is present in an abstract
def isPresent(item, abstract):
    if item in abstract:
        return True
    if '*' in item:
        item = item.replace('*', '.*')
        for word in abstract:
            if re.match(item, word):
                return True
    return False
        
# Excel column number to string. ex. 1 -> A
def colNumToString(div):
    string=""
    while div>0:
        module=(div-1)%26
        string=chr(65+module)+string
        div=int((div-module)/26)
    return string

# Generate an excel file based on results    
def generateExcel(topics, years, results, outputPath):
    wb = Workbook()
    ws = wb.active
    years.sort()
    topics.sort()
    col = 2
    row = 2
    for year in years:
        ws["A" + str(row)] = year
        row += 1
    for topic in topics:
        ws[colNumToString(col) + "1"] = topic
        col += 1
    for result in results:
        row = getRow(ws, result[1], len(years)+2)
        col = getCol(ws, result[0], len(topics)+2)
        ws[colNumToString(col) + str(row)] = result[2]
    wb.save(outputPath)

# Find of given year
def getRow(ws, year, size):
    for row in range(2, size):
        if ws["A" + str(row)].value == year:
            return row
    return -1
    
# Find column of given topic
def getCol(ws, topic, size):
    for col in range(2, size):
        if ws[colNumToString(col) + "1"].value == topic:
            return col
    return -1

# Run everything
def run():
    try:
        updateStatus("Running...")
        abstractPath = abstractEntry.get()
        topicsPath = topicsEntry.get()
        outputPath = outputEntry.get()
        
        topics = getTopics(topicsPath)    
        files = getFiles(abstractPath, [])
        abstracts = getAbstracts(files)
        updateStatus("Analyzing data...")
        results = []
        for topic in topics.keys():
            for year in abstracts.keys():
                topicCount = getTopicCount(topics[topic], abstracts[year])
                if topicCount > 0:
                    results.append((topic, year, topicCount))
                    #print topic + " (" + year + "): " + str(topicCount)
        updateStatus("Writing to excel...")
        generateExcel(topics.keys(), abstracts.keys(), results, outputPath)
        updateStatus("Success. Excel file written to " + outputPath)
    except Exception as e:
        updateStatus("Error: " + str(e))

# Topics dict -> string                
def topicsToString(topics):               
    s = ""
    for topic in topics.keys():
        s = s + topic + ": " + str(topics[topic]) + "\n"
    return s

def updateStatus(status):
    statusLabel.configure(text="Status: " + status)
    w.update_idletasks()
   
def updateTopicsLabel():
    try:
        topicsDict = getTopics(topicsEntry.get())
        topicsList = topicsToString(topicsDict)
    except Exception as e:
        topicsList = "Couldn't load topics. Make sure your path is correct and file is valid."
        print str(e)
    topics.configure(text=topicsList)       

w = Tk()
w.title("Topic frequency analysis")

topicsLabel = Label(w, text="Topics file path (.txt file):")
topicsLabel.grid(row=0, sticky=W, pady=5, padx=5)

topicsEntry = Entry(w, width=50)
topicsEntry.insert(END, 'topics.txt')
topicsEntry.grid(row=0,column=1, sticky=W, pady=5, padx=5)

topicsHeader = Label(w, text="Topics being used:")
topicsHeader.grid(row=1, column=0, sticky=W, pady=5, padx=5)

Button(w, text='Load from file', command=updateTopicsLabel).grid(row=1, column=1, sticky=W, pady=5, padx=5)

topics = Label(w, text="(Loaded topics will appear here)", justify=LEFT)
topics.grid(row=2, sticky=W, padx=20, pady=10, columnspan=2)


abstractLabel = Label(w, text="Abstracts path (folder):")
abstractLabel.grid(row=3, sticky=W, pady=5, padx=5)

abstractEntry = Entry(w, width=50)
abstractEntry.insert(END, 'RS_Abstracts_1975-1936')
abstractEntry.grid(row=3,column=1, sticky=W, pady=5, padx=5)

outputLabel = Label(w, text="Output path (.xlsx):")
outputLabel.grid(row=4, sticky=W, pady=5, padx=5)

outputEntry = Entry(w, width=50)
outputEntry.insert(END, 'output.xlsx')
outputEntry.grid(row=4,column=1, sticky=W, pady=5, padx=5)

Button(w, text='Generate topic frequencies', command=run).grid(row=5, column=0, sticky=E, pady=5, padx=5)
Button(w, text='Quit', command=w.quit).grid(row=5, column=1, sticky=W, pady=5, padx=5)

statusLabel = Label(w, text="Status: idle")
statusLabel.grid(row=6, pady=5, columnspan=2, sticky=W)

w.mainloop()