from docx import Document

from main import color_text


if __name__ == '__main__':
    expected = 'СЛОВО'
    doc = Document('template.docx')
    color_text(doc, expected, 'red')
    doc.save('new.docx')
