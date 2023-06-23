import PyPDF2
import re
import tabula
import langid
from collections import Counter

# Get title
def title_pdf(document):
    with open(document, 'rb') as document:
        reader = PyPDF2.PdfReader(document)
        metadata = reader.metadata
        if metadata is not None:
            title = metadata.get('/Title', '')
            return title
        else:
            return ''

# Count pages 
def page_count_pdf(document):
    with open(document, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        return len(reader.pages)

# Count words
def word_count_pdf(document):
    with open(document, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        words = 0
        for page in reader.pages:
            words += len(page.extract_text().split())
        return words

# Extract keywords  
import re
import PyPDF2
from collections import Counter

def keywords_pdf(document_path):
    with open(document_path, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        text = ""
        for page in reader.pages:
            text += page.extract_text()

    # Remove non-alphanumeric characters and convert to lowercase
    cleaned_text = re.sub(r'\W+', ' ', text.lower())
    
    # Split text into individual words
    words = cleaned_text.split()

    # Count word frequency
    word_count = Counter()

    for word in words:
        if len(word) > 1:
            word_count[word] += 1

    # Extract the top 3 most frequent words that are not single-letter words or numbers
    keywords = []
    for word, count in word_count.most_common():
        if len(keywords) >= 3:
            break
        if len(word) > 1 and not word.isnumeric():
            keywords.append(word)

    return keywords


# Count characters
def character_count_pdf(document):
    with open(document, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        num_characters = 0
        for page in reader.pages:
            text = page.extract_text()
            num_characters += len(text)
        return num_characters
    
# Get typography
def typography_pdf(document):
    with open(document, 'rb') as document:
        reader = PyPDF2.PdfReader(document)
        typography = set()
        for page in reader.pages:
            if '/Font' in page['/Resources']:
                page_fonts = page['/Resources']['/Font']
                for font in page_fonts:
                    if '/BaseFont' in page_fonts[font]:
                        font_name = page_fonts[font]['/BaseFont']
                        typography.add(font_name)
        return list(typography)

# Count images 
def image_count_pdf(document):
    with open(document, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        images = 0
        for page in reader.pages:
            if '/XObject' in page['/Resources']:
                x_objects = page['/Resources']['/XObject']
                for obj in x_objects:
                    if x_objects[obj]['/Subtype'] == '/Image':
                        images += 1
        return images
    
# Count tables
def table_count_pdf(document):
    tables = tabula.read_pdf(document, pages='all', multiple_tables=True, pandas_options={'header': None})
    return len(tables)

# Detect language 
def language_pdf(document):
    with open(document, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        text = ""
        for page in reader.pages:
            text += page.extract_text()
        language = langid.classify(text)[0]
        return language

# Detect encryption
def encryption_pdf(document):
    with open(document, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        if reader.is_encrypted:
            return True
        else:
            return False
        
# Get author
def author_pdf(document):
    with open(document, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        metadata = reader.metadata
        if metadata:
            return metadata.get('/Author', '')
        return ""