import sys
import os
import comtypes.client

path = "docs"
word = comtypes.client.CreateObject("Word.Application")

for filename in os.listdir(path):
    if filename.endswith(".doc") or filename.endswith(".docx"):
        doc_abspath = os.path.abspath(path)
        doc_path = os.path.join(path, filename)
        print(f"Converting {doc_path} to PDF...")
        pdf_path = os.path.join(doc_abspath, filename.replace(".docx", ".pdf"))
        doc = word.Documents.Open(doc_abspath)
        doc.SaveAs(pdf_path, FileFormat=17)
        doc.Close()
        print(f"Converted {doc_path} to PDF.")

word.Quit()

print("All .doc and .docx files in the 'docs' folder have been converted to PDF.")