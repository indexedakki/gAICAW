import zipfile
import os
import tempfile
import shutil
from lxml import etree

def extract_xml_from_docx(docx_path):
    tmp_dir = "C:\\Users\\2113196\\Downloads\\New"
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(tmp_dir)
    xml_path = os.path.join(tmp_dir, 'word/document.xml')
    with open(xml_path, 'r', encoding='utf-8') as file:
        xml_content = file.read()
    return xml_content, tmp_dir

def xml_to_docx(xml_content, output_path):
    tmp_dir = "C:\\Users\\2113196\\Downloads\\New"
    xml_path = os.path.join(tmp_dir, 'word/document.xml')
    with open(xml_path, 'w', encoding='utf-8') as file:
        file.write(xml_content)
    with zipfile.ZipFile(output_path, 'w') as zip_ref:
        for root, dirs, files in os.walk(tmp_dir):
            for file in files:
                zip_ref.write(os.path.join(root, file),
                              os.path.relpath(os.path.join(root, file), tmp_dir))
    # shutil.rmtree(tmp_dir)


tmp_dir = "C:\\Users\\2113196\\Downloads\\new"
xml_path = os.path.join(tmp_dir, 'word/document.xml')
with open(xml_path, 'r', encoding='utf-8') as file:
    xml_content = file.read()

# Extract XML from a Word document
# xml_content, tmp_dir = extract_xml_from_docx("C:\\Users\\2113196\\Downloads\\Table5_1.docx")

print(xml_content)


# Modify the XML content as needed
# For example, you can parse it with lxml and modify it
root = etree.fromstring(xml_content.encode('utf-8'))
# Modify the root element or its children as needed



# Convert the modified XML back to a Word document
xml_to_docx(etree.tostring(root, pretty_print=True).decode('utf-8'), "C:\\Users\\2113196\\Downloads\\Modified_Table5_1.docx")

# Clean up the temporary directory


# import zipfile
# import os
# import tempfile
# import shutil
# from lxml import etree
# def xml_to_docx(xml_content, output_path):
#     tmp_dir = tempfile.mkdtemp()
#     xml_path = os.path.join(tmp_dir, "C:\\Users\\2113196\\Downloads\\Table5_1 xml\\word\\document.xml")
#     with open(xml_path, 'w', encoding='utf-8') as file:
#         file.write(xml_content)
#     with zipfile.ZipFile(output_path, 'w') as zip_ref:
#         for root, dirs, files in os.walk(tmp_dir):
#             for file in files:
#                 zip_ref.write(os.path.join(root, file),
#                               os.path.relpath(os.path.join(root, file), tmp_dir))
#     shutil.rmtree(tmp_dir)

# tmp_dir = tempfile.mkdtemp()
# xml_path = os.path.join(tmp_dir, "C:\\Users\\2113196\\Downloads\\Table5_1 xml\\word\\document.xml")
# with open(xml_path, 'r', encoding='utf-8') as file:
#     xml_content = file.read()
# # Modify the XML content as needed
# # For example, you can parse it with lxml and modify it
# root = etree.fromstring(xml_content.encode('utf-8'))
# # Modify the root element or its children as needed

# # Convert the modified XML back to a Word document
# xml_to_docx(etree.tostring(root, pretty_print=True).decode('utf-8'), "C:\\Users\\2113196\\Downloads\\Modified_Table5_1.docx" )