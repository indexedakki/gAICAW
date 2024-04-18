import tempfile  
import os  
from lxml import etree  
import zipfile  
import shutil  

def write_and_close_docx(zipfile1, xml_content, output_filename):  
    tmp_dir = tempfile.mkdtemp()  
    zipfile1.extractall(tmp_dir)  

    with open(os.path.join(tmp_dir, 'word/document.xml'), 'wb') as f:  
        xmlstr = etree.tostring(xml_content, pretty_print=True)  
        f.write(xmlstr)
    filenames = zipfile1.namelist()  
    with zipfile.ZipFile(output_filename, "w") as docx:  
        for filename in filenames:  
            docx.write(os.path.join(tmp_dir, filename), filename)  
  
    shutil.rmtree(tmp_dir)  

document_path = "C:\\Users\\2113196\\Downloads\\merged.docx" 
output_path = "C:\\Users\\2113196\\Downloads\\merged.docx"  
  
# Initialize the zipfile with the input document path  
zipfile = zipfile.ZipFile(document_path)  
  
# Load the XML content from the input document  
document_xml = zipfile.read('word/document.xml')  
root = etree.fromstring(document_xml)  
  
# Define the namespaces  
namespaces = {  
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",  
}  
  
# Modify the XML content here  
# For example, let's add a new paragraph with text "Hello, World!" at the beginning of the document  
new_paragraph = etree.Element("{%(w)s}p" % namespaces)  
new_run = etree.SubElement(new_paragraph, "{%(w)s}r" % namespaces)  
new_text = etree.SubElement(new_run, "{%(w)s}t" % namespaces)  
new_text.text = "<@word_document_start@>"  
root.insert(0, new_paragraph)
new_paragraph = etree.Element("{%(w)s}p" % namespaces)  
new_run = etree.SubElement(new_paragraph, "{%(w)s}r" % namespaces)  
new_text = etree.SubElement(new_run, "{%(w)s}t" % namespaces)  
new_text.text = "<@word_document_end@>"  
root.append(new_paragraph)

write_and_close_docx(zipfile, root, output_path)