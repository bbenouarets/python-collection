import os
from docx2pdf import convert

path = "docs"

for filename in os.listdir(path):
    if filename.endswith(".doc") or filename.endswith(".docx"):
        doc_path = os.path.join(path, filename)
        # Remove spaces from the filename
        new_doc_path = doc_path.replace(" ", "_")
        # Create empty file with the new name
        open(new_doc_path, "w").close()
        # Rename the file
        convert(doc_path, new_doc_path)

print("All .doc and .docx files in the 'docs' folder have been converted to PDF.")