import re
from docx import Document
import pandas as pd

docx_path = "tekstovye_zadachi_po_matematike.docx"

def parse_doc(docx_path):
    doc = Document(docx_path)
    
    for para in doc.paragraphs:
        print(para.text)

if __name__ == "__main__":
    parse_doc(docx_path=docx_path)