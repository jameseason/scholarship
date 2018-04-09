
def formatDate(prefix):
    return hh.get(prefix + 'm') + '/' + hh.get(prefix + 'd') + '/' + hh.get(prefix + 'y')

### Line 1 - name and occupation
def getLineOne():
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
            s += ' / ' + hh.get('occup' + str(x)):
    s += '\n'
    return s

### Line 2 - address and contact info
def getLineTwo():
    s = hh('address') + ', ' + hh('town') + ', ' + hh('state') + ' ' + hh('zip') + ' ' + hh('telephone')
    # email?
    s += '\n'
    return s
    
### Line 3 - ordinations
def getLineThree():
    s = ''
    if hh.contains('ordain_deac'):
        s += 'Deac. ' + formatDate('ordain_deac') + ';'
    if hh.contains('ordain_mins'):
        s += 'Mins. ' + formatDate('ordain_mins') + ';'
    if hh.contains('ordain_bish'):
        s += 'Bish. ' + formatDate('ordain_bish')
    s += '\n'
    return s

### Line 4 - personal info
def getLineFour():
    s = 'b. ' + formatDate('bornm') + ', '
    if hh.contains('hhhdiedm'):
        s += 'd. ' + formatDate('hhhdied') + ', ' 
    if hh.contains('fatherfirst'):
        s += 's.o. ' + hh.get('fatherfirst') + ' ' + hh.get('fathermiddle') + ' ' + hh.get('fathersuffix')
        if hh.contains('motherfirst'):
            s += ' & ' + hh.get('motherfirst') + ' ' + hh.get('mothermiddle') + ' (' + hh.get('motherlast') + ') ' + hh.get('fatherlast') + ' ' + hh.get('hhparentcode') + ', '  
      
    return s
### Children

### Children of other wives

hh = #todo
s = ''