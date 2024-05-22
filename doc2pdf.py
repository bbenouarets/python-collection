import os  # Importiert das Modul für Betriebssystem-Funktionen
import pypandoc  # Importiert pypandoc zur Umwandlung von .doc-Dateien
import pdfkit  # Importiert pdfkit zur Erstellung von PDFs
from docx import Document  # Importiert die Document-Klasse aus dem python-docx-Modul

# Funktion zur Umwandlung von .docx in .pdf
def convert_docx_to_pdf(docx_path, pdf_path):
    document = Document(docx_path)  # Öffnet das .docx-Dokument
    html_path = docx_path.replace('.docx', '.html')  # Definiert den Pfad für die temporäre HTML-Datei
    document.save(html_path)  # Speichert das Dokument als HTML-Datei
    pdfkit.from_file(html_path, pdf_path)  # Wandelt die HTML-Datei in eine PDF-Datei um
    os.remove(html_path)  # Löscht die temporäre HTML-Datei

# Funktion zur Umwandlung von .doc in .pdf
def convert_doc_to_pdf(doc_path, pdf_path):
    output = pypandoc.convert_file(doc_path, 'pdf', outputfile=pdf_path)  # Wandelt die .doc-Datei direkt in PDF um
    assert output == "", "Conversion failed: {}".format(output)  # Prüft, ob die Umwandlung erfolgreich war

# Funktion zur Umwandlung aller .doc- und .docx-Dateien in einem Ordner in PDFs
def convert_all_docs_to_pdf(folder_path):
    for filename in os.listdir(folder_path):  # Durchläuft alle Dateien im angegebenen Ordner
        file_path = os.path.join(folder_path, filename)  # Erstellt den vollständigen Pfad zur Datei
        if filename.endswith('.docx'):  # Prüft, ob die Datei eine .docx-Datei ist
            pdf_path = file_path.replace('.docx', '.pdf')  # Definiert den Pfad für die PDF-Datei
            convert_docx_to_pdf(file_path, pdf_path)  # Wandelt die .docx-Datei in PDF um
            print(f"Converted {filename} to PDF.")  # Gibt eine Erfolgsmeldung aus
        elif filename.endswith('.doc'):  # Prüft, ob die Datei eine .doc-Datei ist
            pdf_path = file_path.replace('.doc', '.pdf')  # Definiert den Pfad für die PDF-Datei
            convert_doc_to_pdf(file_path, pdf_path)  # Wandelt die .doc-Datei in PDF um
            print(f"Converted {filename} to PDF.")  # Gibt eine Erfolgsmeldung aus

# Fordert den Benutzer auf, den Pfad zum Ordner mit den .doc- und .docx-Dateien einzugeben
folder_path = input("Please enter the path to the folder containing .doc and .docx files: ")
convert_all_docs_to_pdf(folder_path)  # Ruft die Funktion zur Umwandlung aller Dateien auf
