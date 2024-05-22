import os
import win32com.client

path = "docs"

word = win32com.client.Dispatch("Word.Application")

for filename in os.listdir(path):
    if filename.endswith(".doc") or filename.endswith(".docx"):
        doc_path = os.path.join(path, filename)
        print(f"Path: {doc_path}")
        print(f"Converting {doc_path} to PDF...")
        pdf_path = os.path.join(path, filename.replace(".doc", ".pdf").replace(".docx", ".pdf"))
        doc = word.Documents.Open(doc_path)
        doc.SaveAs(pdf_path, FileFormat=17)
        doc.Close()
        print(f"Converted {doc_path} to PDF.")

word.Quit()