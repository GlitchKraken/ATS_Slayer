from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt

doc = Document("./test2.docx")
styles = doc.styles


for style in styles: 
    print(style.name)