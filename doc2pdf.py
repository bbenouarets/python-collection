import os
from docx2pdf import convert

path = "docs"

for filename in os.listdir(path):
    if filename.endswith(".doc") or filename.endswith(".docx"):
        doc_path = os.path.join(path, filename)
        print(f"Path: {doc_path}")
        print(f"Converting {doc_path} to PDF...")
        convert(doc_path)
        print(f"Converted {doc_path} to PDF.")

print("All .doc and .docx files in the 'docs' folder have been converted to PDF.")