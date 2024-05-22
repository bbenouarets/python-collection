import sys
import os
import comtypes.client

path = "docs"
word = comtypes.client.CreateObject("Word.Application")

for filename in os.listdir(path):
    if filename.endswith(".doc") or filename.endswith(".docx"):
        doc_path = os.path.join(path, filename)
        pdf_path = os.path.join(path, filename.replace(".docx", ".pdf").replace(".doc", ".pdf"))
        doc = word.Documents.Open(doc_path)
        doc.SaveAs(pdf_path, FileFormat=17)
        print(f"Converted {filename} to PDF.")
        doc.Close()

word.Quit()

print("All .doc and .docx files in the 'docs' folder have been converted to PDF.")