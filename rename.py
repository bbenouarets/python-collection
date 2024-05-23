# Rename script to replace specific characters in filenames
import os
import re

path = "export"
export = "export_pdf"

replace_search = ".pdfx"
replace_replace = ".pdf"

for filename in os.listdir(path):
    if replace_search in filename:
        new_filename = filename.replace(replace_search, replace_replace)
        os.rename(os.path.join(path, filename), os.path.join(path, new_filename))
        print(f"Renamed {filename} to {new_filename}.")
