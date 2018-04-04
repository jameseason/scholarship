from openpyxl import Workbook, load_workbook
from docx import Document

ATTRS_ROW = '3'

# Load excel file from path
def loadWorkbook(path):
    wb = load_workbook(path)
    ws = wb.active
    return ws
    
# Get an array of dicts of attrs for each households
def getHouseholds(ws):
    numAttrs = getNumAttrs(ws)
    numHouseholds = getNumHouseholds(ws)
    households = []
    for x in range(4, 50):#numHouseholds):
        hh = {}
        for y in range(1, numAttrs):
            attr = ws[colNumToString(y) + ATTRS_ROW].value
            val = ws[colNumToString(y) + str(x)].value
            if not val is None:
                attr = attr.strip().encode('ascii', 'ignore')
                val = val.strip().encode('ascii', 'ignore')
                hh[attr] = val
        households.append(hh)
    return households

# Number of attributes in ws    
def getNumAttrs(ws):
    i = 1
    attr = ws[colNumToString(i) + ATTRS_ROW].value
    while not attr is None:
        i += 1
        attr = ws[colNumToString(i) + ATTRS_ROW].value
    return i

# Number of households in ws    
def getNumHouseholds(ws):    
    i = int(ATTRS_ROW) + 1
    householdHead = ws['I' + str(i)].value
    while not householdHead is None:
        i += 1
        householdHead = ws['I' + str(i)].value
    return i
    
# Convert column number to corresponding letter, ex: 1 -> A, 2 -> B
def colNumToString(div):
    string=""
    while div > 0:
        module=(div-1)%26
        string=chr(65+module)+string
        div=int((div-module)/26)
    return string    
   
# Run everything to get households
def getData():
    ws = loadWorkbook('Groffdale_2017.xlsx')
    households = getHouseholds(ws)
    return households
        
        