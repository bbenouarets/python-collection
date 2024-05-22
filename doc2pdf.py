import sys
import os
import comtypes.client

path = "docs"

word = comtypes.client.CreateObject("Word.Application")

for filename in os.listdir(path):
    if filename.endswith(".doc") or filename.endswith(".docx"):
        doc_path = os.path.join(path, filename)
        print(f"Path: {doc_path}")
        print(f"Converting {doc_path} to PDF...")
        pdf_path = os.path.join(path, filename.replace(".doc", ".pdf").replace(".docx", ".pdf").replace(" ", "_"))
        doc = word.Documents.Open(doc_path)
        doc.SaveAs(pdf_path, FileFormat=17)
        doc.Close()
        print(f"Converted {doc_path} to PDF.")