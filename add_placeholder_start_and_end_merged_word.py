import tempfile  
import os  
from lxml import etree  
import zipfile  
import shutil  

class WordDocumentModifier:
    def __init__(self, docx_path):  
        self.zipfile = zipfile.ZipFile(docx_path)  
  
    def _write_and_close_docx(self, xml_content, output_filename):  
        tmp_dir = tempfile.mkdtemp()  
        self.zipfile.extractall(tmp_dir)  

        with open(os.path.join(tmp_dir, 'word/document.xml'), 'wb') as f:  
            xmlstr = etree.tostring(xml_content, pretty_print=True)  
            f.write(xmlstr)
        filenames = self.zipfile.namelist()  
        with zipfile.ZipFile(output_filename, "w") as docx:  
            for filename in filenames:  
                docx.write(os.path.join(tmp_dir, filename), filename)  
  
        shutil.rmtree(tmp_dir)  

document_path = "C:\\Users\\2113196\\Downloads\\merged.docx" 
output_path = "C:\\Users\\2113196\\Downloads\\merged.docx"  
  
# Initialize the WordDocumentModifier with the input document path  
modifier = WordDocumentModifier(document_path)  
  
# Load the XML content from the input document  
document_xml = modifier.zipfile.read('word/document.xml')  
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

modifier._write_and_close_docx(root, output_path)