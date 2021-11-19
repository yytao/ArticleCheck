from docx import Document


document = Document('./199-235.docx')

for paragraph in document.paragraphs:
    if "所属行业" in paragraph.text:
        paragraph.insert_paragraph_before("第一参赛者姓名：")
        paragraph.insert_paragraph_before("现所在国家/地区：")

document.save('./new_artical.docx')


