from docx import Document

def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None
    #paragraph._p = paragraph._element = None

document = Document('./all.docx')

for paragraph in document.paragraphs:
    if "第一参赛者姓名" in paragraph.text:
        delete_paragraph(paragraph)

    if "现所在国家" in paragraph.text:
        delete_paragraph(paragraph)



document.save('./result.docx')
