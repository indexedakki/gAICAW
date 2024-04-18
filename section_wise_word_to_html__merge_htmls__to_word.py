import os
import pypandoc
from docx import Document

from bs4 import BeautifulSoup

# Convert Word to HTML
def word_to_html(word_file, html_name):
    output = pypandoc.convert_file(word_file, 'html', outputfile=f"C:\\Users\\2113196\\Downloads\\html_split\\{html_name}")
    return output

def merge_html_files(html_dir, html_files):  
    # create a new HTML file to merge the contents into
    with open("C:\\Users\\2113196\\Downloads\\merged.html", 'w', encoding="utf-8") as outfile:
        # iterate over each HTML file in the directory
        for html_file in html_files:
            # construct the full path to the HTML file
            html_path = os.path.join(html_dir, html_file)
            
            # iterate over each line in the HTML file
            for line in open(html_path, encoding="utf-8"):
                # write the line to the output file
                outfile.write(line)

# Convert HTML to Word
def html_to_word(html_file, merged_docx_file_path):
    output = pypandoc.convert_file(html_file, 'docx', outputfile=merged_docx_file_path, encoding="utf-8")
    return output

def run_functions():
    # directory = "C:\\Users\\2113196\\Downloads\\word_split\\"
    # sections_present_in_word = len([f for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f))])
    # for i in range(1, sections_present_in_word+1):
    #     doc_name = f"section {i}.docx"
    #     html_name = f"section {i}.html"
    #     print(html_name)
    #     word_to_html(f"{directory}{doc_name}", html_name)

    # html_dir = "C:\\Users\\2113196\\Downloads\\html_split\\"
    # sections_present_in_html = len([f for f in os.listdir(html_dir) if os.path.isfile(os.path.join(html_dir, f))])
    # html_files = []
    # for i in range(1, sections_present_in_html+1):
    #     html_name = f"section {i}.html"
    #     html_files.append(html_name)
    # merge_html_files(html_dir, html_files)

    merged_html_file_path = "C:\\Users\\2113196\\Downloads\\merged.html"
    merged_docx_file_path = "C:\\Users\\2113196\\Downloads\\merged.docx"
    html_to_word(merged_html_file_path, merged_docx_file_path)
run_functions()