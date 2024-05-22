import sys
import os
import comtypes.client

path = "docs"

word = comtypes.client.CreateObject("Word.Application")

for filename in os.listdir(path):
    if filename.endswith(".doc") or filename.endswith(".docx"):
        doc_path = os.path.join(path, filename)
        abs_doc_path = os.path.abspath(doc_path)  # Get absolute path
        print(f"Path: {abs_doc_path}")
        print(f"Converting {abs_doc_path} to PDF...")
        pdf_path = os.path.join(path, filename.replace(".doc", ".pdf").replace(".docx", ".pdf").replace(" ", "_"))
        abs_pdf_path = os.path.abspath(pdf_path)  # Get absolute path
        doc = word.Documents.Open(abs_doc_path)
        doc.SaveAs(abs_pdf_path, FileFormat=17)
        doc.Close()
        print(f"Converted {abs_doc_path} to PDF.")