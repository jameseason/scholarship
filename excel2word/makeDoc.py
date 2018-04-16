from extractExcel import getData
from prepHousehold import getLineOne, getLineTwo, getLineThree, getLineFour, getChildren
from docx import Document
from docx.shared import Inches, Cm, Pt
from docx.enum.section import WD_SECTION

from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def exampleDoc():
    document = Document()

    document.add_heading('Document Title', 0)
    
    p = document.add_paragraph('A plain paragraph having some ')
    p.add_run('bold').bold = True
    p.add_run(' and some ')
    p.add_run('italic.').italic = True

    document.add_heading('Heading, level 1', level=1)
    document.add_paragraph('Intense quote', style='Intense Quote')

    document.add_paragraph(
        'first item in unordered list', style='List Bullet'
    )
    document.add_paragraph(
        'first item in ordered list', style='List Number'
    )

    table = document.add_table(rows=1, cols=3)
   
    tc = table.cell(0,0)._tc     # As a test, fit text to cell 0,0
    tcPr = tc.get_or_add_tcPr()

    tcFitText = OxmlElement('w:tcLeftPadding')
    tcFitText.set(qn('w:val'),"0")
    tcPr.append(tcFitText)
    
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Qty'
    hdr_cells[1].text = 'Id'
    hdr_cells[2].text = 'Desc'
    document.add_page_break()

    document.save('demo.docx')
    
def testDoc(hhs):
    doc = Document()
    doc.add_paragraph("first page blank.")
    doc.add_page_break()
    section = doc.add_section(WD_SECTION.NEW_PAGE)
    set_number_of_columns(section, 2)
    style = doc.styles['Normal']
    
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(10)

    paragraph_format = style.paragraph_format
    paragraph_format.space_before = 0
    paragraph_format.space_after = 0
    
    for hh in hhs:
        doc = getLineOne(hh, doc)
        two = doc.add_paragraph(getLineTwo(hh).strip())
        three = getLineThree(hh).strip()
        if len(three) > 0:
            doc.add_paragraph(three)
        four = doc.add_paragraph(getLineFour(hh))
        doc = getChildren(doc, hh)
        doc.add_paragraph("")
    
    secs = doc.sections
    for sec in secs:
        sec.top_margin = Cm(.5)
        sec.bottom_margin = Cm(.5)
        sec.left_margin = Cm(.5)
        sec.right_margin = Cm(.5)
    doc.save('test.docx')
   

def set_number_of_columns(section, cols):
    """ sets number of columns through xpath. """
    WNS_COLS_NUM = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}num"
    section._sectPr.xpath("./w:cols")[0].set(WNS_COLS_NUM, str(cols))
    


#exampleDoc()    
#quit()

print 'started'
hhs = getData()
print 'got data.'
testDoc(hhs)
print 'done'
