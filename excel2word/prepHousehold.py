'''
1 - occupation: right justified
2 - select which elements to include
3 - adjust font size
(send groff code)
'''



from extractExcel import getData

def formatDate(prefix, hh):
    return hh.get(prefix + 'm') + '/' + hh.get(prefix + 'd') + '/' + hh.get(prefix + 'y')

### Line 1 - name and occupation
def getLineOne(hh):
    s = hh.get('hhcode') + ' ' + hh.get('lastname') + ', ' + hh.get('firstname') + ' ' + hh.get('middle')
    if hh.contains('suffix'):
        s += ', ' + hh.get('suffix')
    if hh.contains('member'):
        if hh.get('member') == 'y':
            s += '*'
    if hh.contains('occup1'):
        s += ' ' + hh.get('occup1')
    for x in range(2,5):
        if hh.contains('occup' + str(x)):
            s += ' / ' + hh.get('occup' + str(x))
    return s

### Line 2 - address and contact info
def getLineTwo(hh):
    s = hh.get('address') + ', ' + hh.get('town') + ', ' + hh.get('state') + ' ' + hh.get('zip') + ' ' + hh.get('telephone')
    # email?
    return s
    
### Line 3 - ordinations
def getLineThree(hh):
    s = ''
    if hh.contains('ordain_deac'):
        s += 'Deac. ' + formatDate('ordain_deac', hh) + ';'
    if hh.contains('ordain_mins'):
        s += 'Mins. ' + formatDate('ordain_mins', hh) + ';'
    if hh.contains('ordain_bish'):
        s += 'Bish. ' + formatDate('ordain_bish', hh)
    return s

### Line 4 - personal info
def getLineFour(hh):
    s = 'b. ' + formatDate('born', hh) + ', '
    if hh.contains('hhhdiedm'):
        s += 'd. ' + formatDate('hhhdied', hh) + ', ' 
    if hh.contains('fatherfirst'):
        s += 's.o. ' + hh.get('fatherfirst') + ' ' + hh.get('fathermiddle') + ' ' + hh.get('fathersuffix')
        if hh.contains('motherfirst'):
            s += ' & ' + hh.get('motherfirst') + ' ' + hh.get('mothermiddle') + ' (' + hh.get('motherlast') + ') ' + hh.get('fatherlast') + ' ' + hh.get('hhhparentcode')
    if hh.contains('hhhmovedfrom'):
        s += ', moved in ' + formatDate('hhhmovedfrom', hh) + ' from ' + hh.get('hhhmovedfrom')
    s += '; '
    if hh.contains('cmary'):
        s += 'm. ' + formatDate('cmar', hh)
    if hh.contains('cwfname'):
        s += ' to ' + hh.get('cwfname') + ' ' + hh.get('cwmiddle') + ' ' + hh.get('cwlname')
        if hh.get('cwmember') == 'y':
            s += '*'
        s += ', b. ' + formatDate('cwborn', hh)
        if hh.contains('cwdiedy'):
            s += ', d. ' + formatDate('cwdied', hh)
        if hh.contains('cwdadfname'):
            s += ', d.o. ' + hh.get('cwdadfname') + ' ' + hh.get('cwdadmname') + ' ' + hh.get('cwdadlname') + ' ' + hh.get('cwdadsname')
            if hh.contains('cwmomfname'):
                s += ' & ' + hh.get('cwmomfname') + ' ' + hh.get('cwmommname') + ' ' + hh.get('cwmomlname')
            s += ' ' + hh.get('cwparentcode')
    if hh.contains('cwmovedfrom'):
        s += ', moved in ' + formatDate('cwmovedfrom', hh) + ' from ' + hh.get('cwmovedfrom') + '.'
    if hh.contains('movedfrom'):
        s += ' Moved in ' + formatDate('movedfrom', hh) + ' from ' + hh.get('movedfrom') + '.'
    return s
    
### Children
def getChildren(doc, hh):
    ## First middle last suff | bday | code (or d.) | wifef,m,l:addr code
    table = doc.add_table(rows=1, cols=4)
    #table.autofit = True
    table.style = None
    cells = table.rows[0].cells
    n = 1
    while hh.contains('cwc' + str(n).zfill(2) + 'fname'):
        cells[0].text = hh.get('cwc' + str(n).zfill(2) + 'fname') + ' ' + hh.get('cwc' + str(n).zfill(2) + 'mname') + ' ' + hh.get('cwc' + str(n).zfill(2) + 'sname')
        cells[1].text = formatDate('cwc' + str(n).zfill(2) + 'born', hh)
        if hh.contains('cwc' + str(n).zfill(2) + 'diedy'):
            cells[2].text = 'd.'
            cells[3].text = formatDate('cwc' + str(n).zfill(2) + 'died', hh)
        else:
            cells[2].text = hh.get('cwc' + str(n).zfill(2) + 'bc')
            s = hh.get('cwc' + str(n).zfill(2) + 'spousefname') + ' ' + hh.get('cwc' + str(n).zfill(2) + 'spousemname') + ' ' + hh.get('cwc' + str(n).zfill(2) + 'spouselname') + ' ' + hh.get('cwc' + str(n).zfill(2) + 'spousesname')
            if len(s.strip()) > 0 and hh.contains('cwc' + str(n).zfill(2) + 'address'):
                s += ':'
            s += hh.get('cwc' + str(n).zfill(2) + 'address') + ' '
            s += ' ' + hh.get('cwc' + str(n).zfill(2) + 'hshld#')
            cells[3].text = s.strip()
        n += 1
        if hh.contains('cwc' + str(n).zfill(2) + 'fname'):
            cells = table.add_row().cells
            
    for column in table.columns:
        for cell in column.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcW = tcPr.get_or_add_tcW()
            tcW.type = 'auto'
    return doc
        
    

### Children of other wives

'''
hhs = getData()
for hh in hhs:
    print getLineOne(hh)
    print getLineTwo(hh)
    print getLineThree(hh)
    print getLineFour(hh)
    print '--'
    '''
    