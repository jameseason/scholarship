'''
1 - occupation: right justified
https://stackoverflow.com/questions/28884114/python-docx-align-both-left-and-right-on-same-line?noredirect=1&lq=1

2 - select which elements to include
3 - adjust font size
(send groff code)

table style: https://github.com/python-openxml/python-docx/issues/9

'''

from extractExcel import getData
from docx.shared import Inches, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def formatDate(prefix, hh, z=False):
    if z:
        return hh.get(prefix + 'm').zfill(2) + '/' + hh.get(prefix + 'd').zfill(2) + '/' + hh.get(prefix + 'y').zfill(2)
    else:
        return hh.get(prefix + 'm') + '/' + hh.get(prefix + 'd') + '/' + hh.get(prefix + 'y')
        

### Line 1 - name and occupation
def getLineOne(hh, doc):
    p = doc.add_paragraph('')

    s = hh.get('hhcode') + ' ' + hh.get('lastname').upper() + ', ' + hh.get('firstname').upper() + ' ' + hh.get('middle').upper()
    if hh.contains('suffix'):
        s += ', ' + hh.get('suffix').upper()
    if hh.contains('member'):
        if hh.get('member') == 'y':
            s += '*'
    p.add_run(s.strip()).bold = True
    
    if hh.contains('occup1'):
        s = ' ' + hh.get('occup1').title()
    for x in range(2,5):
        if hh.contains('occup' + str(x)):
            s += ' / ' + hh.get('occup' + str(x)).title()
    p.add_run(s)
    
    return doc

### Line 2 - address and contact info
def getLineTwo(hh):
    s = hh.get('address').title() + ', ' + hh.get('town').title() + ', ' + hh.get('state').upper() + ' ' + hh.get('zip') + ' ' + hh.get('telephone')
    # email?
    return s
    
### Line 3 - ordinations
def getLineThree(hh):
    s = ''
    if hh.contains('ordain_deac'):
        s += 'Deac. ' + formatDate('ordain_deac', hh) + '; '
    if hh.contains('ordain_mins'):
        s += 'Mins. ' + formatDate('ordain_mins', hh) + '; '
    if hh.contains('ordain_bish'):
        s += 'Bish. ' + formatDate('ordain_bish', hh)
    return s

### Line 4 - personal info
def getLineFour(hh):
    s = 'b. ' + formatDate('born', hh) + ', '
    if hh.contains('hhhdiedm'):
        s += 'd. ' + formatDate('hhhdied', hh) + ', ' 
    if hh.contains('fatherfirst'):
        s += 's.o. ' + hh.get('fatherfirst').title() + ' ' + hh.get('fathermiddle').title() + ' ' + hh.get('fathersuffix').title()
        if hh.contains('motherfirst'):
            s += ' & ' + hh.get('motherfirst').title() + ' ' + hh.get('mothermiddle').title() + ' (' + hh.get('motherlast').title() + ') ' + hh.get('fatherlast').title()
        if hh.contains('hhhparentcode'):
            s += ' [' + hh.get('hhhparentcode') + ']'
    if hh.contains('hhhmovedfrom'):
        s += ', moved in ' + formatDate('hhhmovedfrom', hh) + ' from ' + hh.get('hhhmovedfrom').title()
    s += '; '
    if hh.contains('cmary'):
        s += 'm. ' + formatDate('cmar', hh)
    if hh.contains('cwfname'):
        s += ' to ' + hh.get('cwfname').upper() + ' ' + hh.get('cwmiddle').upper() + ' ' + hh.get('cwlname').upper()
        if hh.get('cwmember') == 'y':
            s += '*'
        s += ', b. ' + formatDate('cwborn', hh)
        if hh.contains('cwdiedy'):
            s += ', d. ' + formatDate('cwdied', hh)
        if hh.contains('cwdadfname'):
            s += ', d.o. ' + hh.get('cwdadfname').title() + ' ' + hh.get('cwdadmname').title() + ' ' + hh.get('cwdadlname').title() + ' ' + hh.get('cwdadsname').title()
            if hh.contains('cwmomfname'):
                s += ' & ' + hh.get('cwmomfname').title() + ' ' + hh.get('cwmommname').title() + ' (' + hh.get('cwmomlname').title() + ') ' + hh.get('cwdadlname').title()
            if hh.contains('cwparentcode'):
                s += ' [' + hh.get('cwparentcode') + ']'
    if hh.contains('cwmovedfrom'):
        s += ', moved in ' + formatDate('cwmovedfrom', hh) + ' from ' + hh.get('cwmovedfrom').title() + '.'
    if hh.contains('movedfrom'):
        s += ' Moved in ' + formatDate('movedfrom', hh) + ' from ' + hh.get('movedfrom').title() + '.'
    return s
    
### Children
def getChildren(doc, hh):
    if not hh.contains('cwc01fname'):
        return doc
    ## First middle last suff | bday | code (or d.) | wifef,m,l:addr code
    table = doc.add_table(rows=1, cols=4)
    table = set_col_widths(table)
    table.style = 'Table Grid'
    
    cells = table.rows[0].cells

    n = 1
    while hh.contains('cwc' + str(n).zfill(2) + 'fname'):
        name = hh.get('cwc' + str(n).zfill(2) + 'fname').title() + ' '
        name += hh.get('cwc' + str(n).zfill(2) + 'mname').title() + ' ' 
        name += hh.get('cwc' + str(n).zfill(2) + 'sname').title()
        cells[0].text = name
        cells[1].text = formatDate('cwc' + str(n).zfill(2) + 'born', hh, True)
        if hh.contains('cwc' + str(n).zfill(2) + 'diedy'):
            cells[2].text = 'd.'
            cells[3].text = formatDate('cwc' + str(n).zfill(2) + 'died', hh, True)
        else:
            cells[2].text = hh.get('cwc' + str(n).zfill(2) + 'bc').upper()
            s = hh.get('cwc' + str(n).zfill(2) + 'spousefname').title() + ' ' 
            s += hh.get('cwc' + str(n).zfill(2) + 'spousemname').title() + ' ' 
            s += hh.get('cwc' + str(n).zfill(2) + 'spouselname').title() + ' ' 
            s += hh.get('cwc' + str(n).zfill(2) + 'spousesname').title()
            if len(s.strip()) > 0 and hh.contains('cwc' + str(n).zfill(2) + 'address'):
                s += ':'
            s += hh.get('cwc' + str(n).zfill(2) + 'address').title() + ' '
            if hh.contains('cwc' + str(n).zfill(2) + 'hshld#'):
                s += ' [' + hh.get('cwc' + str(n).zfill(2) + 'hshld#') + ']'
            cells[3].text = s.strip()
        n += 1
        if hh.contains('cwc' + str(n).zfill(2) + 'fname'):
            cells = table.add_row().cells
            
    table = set_font_size(table, 8)
    
    return doc
        
# Set column widths
def set_col_widths(table):
    widths = (1, .7, .3, 1.65)
    i=0
    for cell in table.columns[0].cells:
        cell.width = Inches(widths[i])
        #TODO: remove padding
        # https://groups.google.com/forum/#!topic/python-docx/ABjObkKkOu0
        i += 1
    i = 0
    for column in table.columns:
        column.width = Inches(widths[i])
        i += 1
    return table

# Set font size for a table
def set_font_size(table, size):
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    font = run.font
                    font.size= Pt(size)  
    
### Children of other wives

//todo