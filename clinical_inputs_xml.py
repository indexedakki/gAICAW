import zipfile
import os
import tempfile
import shutil
from lxml import etree

import pandas as pd
import xml.etree.ElementTree as ET

def extract_xml_from_docx(docx_path):
    tmp_dir = "C:\\Users\\2113196\\Downloads\\new"
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(tmp_dir)
    xml_path = os.path.join(tmp_dir, 'word/document.xml')
    with open(xml_path, 'r', encoding='utf-8') as file:
        xml_content = file.read()
    return xml_content, tmp_dir

xml_content, tmp_dir = extract_xml_from_docx("C:\\Users\\2113196\\Downloads\\Table5_1.docx")


def creation_subject_exposure_clinical_inputs_table(file_path):
    skip_rows = 4
    df_clinical_inputs = pd.read_excel(file_path, skiprows=skip_rows )
    patients_treated_sum = df_clinical_inputs['Patients treated'].sum()

    completed_studies_sum = df_clinical_inputs.loc[df_clinical_inputs['Status'] == 'Completed', 'Patients treated'].sum()
    ongoing_studies_sum = df_clinical_inputs.loc[df_clinical_inputs['Status'] == 'Ongoing', 'Patients treated'].sum()

    print(patients_treated_sum)
    print(completed_studies_sum)

    xml_file_path = 'C:\\Users\\2113196\\Downloads\\New\\word\\document.xml'

    # Load the XML file
    tree = ET.parse(xml_file_path)
    root = tree.getroot()

    # Replace 'AAA' with the sum
    for elem in root.iter():
        if elem.text:
            elem.text = elem.text.replace('AAA', str(completed_studies_sum))
            elem.text = elem.text.replace('AAB', str(ongoing_studies_sum))
            elem.text = elem.text.replace('AAC', str(patients_treated_sum))
    tree.write(xml_file_path)

file_path = 'C:\\Users\\2113196\\Downloads\\Clinical inputs_Section 5.1.xlsx'
creation_subject_exposure_clinical_inputs_table(file_path)
    

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
                
# # Extract XML from a Word document
tmp_dir = "C:\\Users\\2113196\\Downloads\\new"
xml_path = os.path.join(tmp_dir, 'word\\document.xml')
with open(xml_path, 'r', encoding='utf-8') as file:
    xml_content = file.read()
                
# Modify the XML content as needed
# For example, you can parse it with lxml and modify it
root = etree.fromstring(xml_content.encode('utf-8'))
# Modify the root element or its children as needed

# Convert the modified XML back to a Word document
xml_to_docx(etree.tostring(root, pretty_print=True).decode('utf-8'), "C:\\Users\\2113196\\Downloads\\Modified_Table5_1.docx")
