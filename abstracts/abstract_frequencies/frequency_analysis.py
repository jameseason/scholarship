# James Eason

import os
import sys
from openpyxl import Workbook
from Tkinter import *
import Tkinter, Tkconstants, tkFileDialog

### Abstract Retrieval ###

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

### Excel Writing ###

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
    
# Excel column number to string. ex. 1 -> A
def colNumToString(div):
    string=""
    while div>0:
        module=(div-1)%26
        string=chr(65+module)+string
        div=int((div-module)/26)
    return string

### Topics ###
    
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
    
# Topics dict -> string                
def topicsToString(topics):               
    s = ""
    i = list(topics.keys())
    i.sort()
    for topic in i:
        s = s + topic + ": " + str(topics[topic]) + "\n"
    return s

### Main Program and GUI ###

# Run everything
def run():
    try:
        updateStatus("Running...")
        abstractPath = abstractEntry.get()
        outputPath = outputEntry.get()
            
        topics = availableTopics
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

# Update status label at bottom of main window
def updateStatus(status):
    statusLabel.configure(text="Status: " + status)
    w.update_idletasks()
   
def addTopicsFromPath(topicsPath):
    try:
        topicsDict = getTopics(topicsPath)
        for topic in topicsDict:
            availableTopics[topic] = topicsDict[topic]
    except Exception as e:
        updateStatus("Couldn't load topics. Make sure your path is correct and file is valid.")
        print str(e)
    refreshTopics()    

# Import topics
def openFileDialog():
    w.filename = tkFileDialog.askopenfilename(initialdir = os.getcwd(),title = "Select topics file",filetypes = (("text files","*.txt"),("all files","*.*")))
    if w.filename != None:
        addTopicsFromPath(w.filename)

# Export topics
def saveFileDialog():
    w.filename = tkFileDialog.asksaveasfilename(initialdir = os.getcwd(),title = "Select topics file",filetypes = (("text files","*.txt"),("all files","*.*")))
    if w.filename != None:
        writeFile(w.filename)        

# Write exported topics file
def writeFile(filepath):
    if not '.txt' in filepath:
        filepath += '.txt'
    file = open(filepath, 'w')
    for topic in availableTopics.keys():
        s = topic + " ("
        for item in availableTopics[topic]:
            s += item + ", "
        s = s[0:-2] + ")\n"
        file.write(s)
    file.close()

# Refresh list of topics
def refreshTopics():
    topicsList = topicsToString(availableTopics)
    topics.configure(text=topicsList)

# Add topic to list from 'New Topic' window
def addTopic(title, items, t):
    title = title.get()
    items = items.get("1.0",'end-1c')
    items = items.split(',')
    for x in range(len(items)):
        items[x] = items[x].strip().encode('ascii','ignore')
        if len(items[x]) < 1:
            items.remove(items[x])
    availableTopics[title] = items 
    refreshTopics()
    t.destroy()
    
    
# Create topics
def newTopicWindow():
    t = Toplevel(w)
    t.wm_title("New topic")
    titleLabel = Label(t, text="Topic title: ")
    titleLabel.grid(row=0, column=0, sticky=W, pady=5, padx=5)
    titleEntry = Entry(t, width=50)
    titleEntry.grid(row=0, column=1, sticky=W, pady=5, padx=5)
    topicsLabel = Label(t, text="List comma separated topic terms. Use an asterisk to match chunks of words.")
    topicsLabel.grid(row=1, sticky=W, pady=5, padx=5, columnspan=2)
    topicsEntry = Text(t, height=5)
    topicsEntry.grid(row=2, sticky=W, pady=5, padx=5, columnspan=2)
    Button(t, text='Save', command=lambda: addTopic(titleEntry, topicsEntry, t)).grid(row=3, column=0, sticky=E, pady=5, padx=5)
    Button(t, text='Quit', command=t.destroy).grid(row=3, column=1, sticky=W, pady=5, padx=5)
    

availableTopics = {}  

w = Tk()
w.title("Topic frequency analysis")

Button(w, text='New topic', command=newTopicWindow).grid(row=0, column=0, sticky=W, pady=5, padx=5)
Button(w, text='Import topics', command=openFileDialog).grid(row=0, column=1, sticky=W, pady=5, padx=5)
Button(w, text='Export topics', command=saveFileDialog).grid(row=0, column=2, sticky=W, pady=5, padx=5)

topics = Label(w, text="(No topics present)", justify=LEFT)
topics.grid(row=1, sticky=W, padx=20, pady=10, columnspan=2)


"""
topicsLabel = Label(w, text="Topics file path (.txt file):")
topicsLabel.grid(row=0, sticky=W, pady=5, padx=5)

topicsEntry = Entry(w, width=50)
topicsEntry.insert(END, 'topics.txt')
topicsEntry.grid(row=0,column=1, sticky=W, pady=5, padx=5)

topicsHeader = Label(w, text="Topics being used:")
topicsHeader.grid(row=1, column=0, sticky=W, pady=5, padx=5)

Button(w, text='Load from file', command=openFileDialog).grid(row=1, column=1, sticky=W, pady=5, padx=5)
"""



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