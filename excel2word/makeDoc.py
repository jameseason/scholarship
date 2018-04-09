from extractExcel import getData
from docx import Document
from docx.shared import Inches

from docx.enum.section import WD_SECTION

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
    
    s = '' #here
    
    doc.add_paragraph(s)
    
    doc.save('test.docx')
   

def set_number_of_columns(section, cols):
    """ sets number of columns through xpath. """
    WNS_COLS_NUM = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}num"
    section._sectPr.xpath("./w:cols")[0].set(WNS_COLS_NUM, str(cols))

print 'started'
hhs = getData()
print 'got data.'
testDoc(hhs)
print 'done'