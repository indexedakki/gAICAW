import zipfile
import os
import tempfile
import shutil
from lxml import etree

import pandas as pd
import xml.etree.ElementTree as ET

def xml_replacment():
    # Load the XML file
    tree = ET.parse(file_path)
    root = tree.getroot()

    print(root)
    # Display the content of the XML file
    # for child in root:
    #     print(ET.tostring(child, encoding='utf8').decode('utf8'))

    # Replace text 'AAA' with '69'
    for elem in root.iter():
        if elem.text:
            elem.text = elem.text.replace('45', '')
        if elem.tail:
            elem.tail = elem.tail.replace('or', '')
        if elem.text:
            elem.text = elem.text.replace('##', '67')
        if elem.tail:
            elem.tail = elem.tail.replace('AAC', '79')

    # Save the modified XML file
    tree.write(file_path)

file_path = "C:\\Users\\2113196\\Downloads\\New\\word\\document.xml"
xml_replacment(file_path)