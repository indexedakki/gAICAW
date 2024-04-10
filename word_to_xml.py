import zipfile
import os
def extract_xml_from_docx(docx_path):
    tmp_dir = "C:\\Users\\2113196\\Downloads\\Table5_1"
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(tmp_dir)
    xml_path = os.path.join(tmp_dir, 'word/document.xml')
    with open(xml_path, 'r', encoding='utf-8') as file:
        xml_content = file.read()
    return xml_content, tmp_dir

# Extract XML from a Word document
# tmp_dir = "C:\\Users\\2113196\\Downloads\\new"
# xml_path = os.path.join(tmp_dir, 'word/document.xml')
# print(xml_path)
# with open(xml_path, 'r', encoding='utf-8') as file:
#     xml_content = file.read()

xml_content, tmp_dir = extract_xml_from_docx("C:\\Users\\2113196\\Downloads\\Table5_1.docx")