import os
import platform
import PyPDF2
import json
import csv
import datetime
from docx import Document
# PDF functions
from PDF import *
# DOCX functions
from DOCX import *

# Document Detective
def Detective(document, document_type, generate_csv, generate_json, generate_html):
    # Get document name
    document_name = os.path.basename(document)
    # Get document path
    document_path = os.path.abspath(document)
    # Get document size
    size = os.path.getsize(document)

    # PDF
    if document_type == 'pdf':
        document_name = os.path.basename(document)
        title = title_pdf(document)
        pages = page_count_pdf(document)
        words = word_count_pdf(document)
        keywords = keywords_pdf(document)
        language = language_pdf(document)
        characters = character_count_pdf(document)
        typography = typography_pdf(document)
        images = image_count_pdf(document)
        tables = table_count_pdf(document)
        encryption = encryption_pdf(document)
        author = author_pdf(document)
        text = ""
        with open(document, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            for page in reader.pages:
                text += page.extract_text()
    # DOCX
    elif document_type == 'docx':
        document_name = os.path.basename(document)
        title = title_docx(document)
        words = word_count_docx(document)
        pages = page_count_docx(document)
        language = language_docx(document)
        characters = character_count_docx(document)
        images = image_count_docx(document)
        typography = typography_docx(document)
        tables = table_count_docx(document)
        keywords = keywords_docx(document)
        encryption = encryption_docx(document)
        author = author_docx(document)
        doc = Document(document)
        text = ""
        for paragraph in doc.paragraphs:
            text += paragraph.text
    else:
        print("unsupported document type - use a PDF or DOCX file.")
        return

    # Get document creation date
    creation_date = datetime.datetime.fromtimestamp(os.path.getctime(document)).strftime('%Y-%m-%d %H:%M:%S')
    # Get document modification date
    modified_date = datetime.datetime.fromtimestamp(os.path.getmtime(document)).strftime('%Y-%m-%d %H:%M:%S')

    # Read HTML template
    with open('lib/template.html', 'r') as html_document:
        html_template = html_document.read()

    # HTML placeholders with dynamic values
    html_content = html_template.format(
        name=document_name,
        path=document,
        type=document_type,
        size=size,
        title=title,
        pages=pages,
        words=words,
        keywords=keywords,
        characters=characters,
        typography=typography,
        images=images,
        tables=tables,
        language=language,
        creation_date=creation_date,
        modified_date=modified_date,
        encryption=encryption,
        author=author
    )

    # Write HTML content to document
    if generate_html:
        with open('reports/report.html', 'w') as html_document:
            html_document.write(html_content)
        print("HTML Report generated in '/reports/report.html'")

    # Generate JSON report
    json_data = {}
    if generate_json:
        json_data = {
            "document": {
                "name": document_name,
                "path": document_path,
                "type": document_type,
                "size": size,
                "title": title,
                "page_count": pages,
                "word_count": words,
                "keywords": keywords,
                "character_count": characters,
                "typography": typography,
                "image_count": images,
                "table_count": tables,
                "language": language,
                "creation_date": creation_date,
                "modified_date": modified_date,
                "encryption": encryption,
                "author": author
            }
        }

        # Write JSON report to document
        json_report = json.dumps(json_data, indent=4)
        with open('reports/report.json', 'w') as json_document:
            json_document.write(json_report)
        print("JSON Report generated in '/reports/report.json'")

    # Generate CSV report
    csv_data = []
    if generate_csv:
        csv_data = [
            ["Name", "Path", "Type", "Size", "Title", "Page count", "Word count", "Keywords", "Character count", "Typography", "Image count", "Table count", "Language", "Creation date", "Modified date", "Encryption", "Author"],
            [document_name, document_path, document_type, size, title, pages, words, keywords, characters, typography, images, tables, language, creation_date, modified_date, encryption, author]
        ]

        with open('reports/report.csv', 'w', newline='') as csv_document:
            writer = csv.writer(csv_document)
            writer.writerows(csv_data)
        print("CSV Report generated in '/reports/report.csv'")

# Clear terminal
def clear_terminal():
    system = platform.system()
    if system == 'Windows':
        os.system('cls')
    else:
        os.system('clear')

# Document Report
class Report:
    def __init__(self):
        self.init_ui()

    def init_ui(self):
        clear_terminal()
        document = input("Enter the relative document path (e.g., documents/document.pdf): ")
        document_type = input("Enter the document type (pdf/docx): ").lower()
        generate_csv = input("Generate CSV report? (y/n): ").lower() == 'y'
        generate_json = input("Generate JSON report? (y/n): ").lower() == 'y'
        generate_html = input("Generate HTML report? (y/n): ").lower() == 'y'

        Detective(document, document_type, generate_csv, generate_json, generate_html)

document_Detective = Report()