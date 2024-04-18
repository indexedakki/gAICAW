namespaces = {  
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",  
}

def word_document_to_xml(docx_path):  
    with zipfile.ZipFile(docx_path) as zf:  
        document_xml = zf.read('word/document.xml')  
    return etree.fromstring(document_xml)  


def extract_text_between_strings(source_root, text_before, text_after, namespaces):  
    text_nodes = source_root.findall(".//{%(w)s}t" % namespaces)  
    extracted_text = []  
    is_collecting = False  
  
    for node in text_nodes:  
        if node.text == text_before:  
            is_collecting = True  
            continue  
        if node.text == text_after:  
            is_collecting = False  
            continue  
        if is_collecting:  
            extracted_text.append(node.text)  
  
    return " ".join(extracted_text)  


source_document_path = "RSI_Sevelamer 800 mg film-coated tablets.docx"
source_root = word_document_to_xml(source_document_path)

# Define the text before and after the text you want to extract
text_before = "Mechanism of action"
text_after = "Pharmacodynamic effect"

# Extract the text between the given strings in the source XML content
extracted_text = extract_text_between_strings(source_root, text_before, text_after, namespaces)
print("Extracted text:", extracted_text)