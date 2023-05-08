import os
import pandas as pd
from docx import Document

# Read Excel data
excel_file = 'data.xlsx'
data = pd.read_excel(excel_file)

# Load Word template
template_file = 'template.docx'

# Output directory for generated documents
output_dir = 'output'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

def replace_placeholders(paragraph, record):
    """Replace placeholders in a paragraph with values from a record."""
    for key, value in record.items():
        placeholder = f'{{{{{key}}}}}'
        if placeholder in paragraph.text:
            paragraph.text = paragraph.text.replace(placeholder, str(value))

# Perform mail merge
for index, row in data.iterrows():
    # Load the Word template for each record
    doc = Document(template_file)

    # Replace placeholders in the document with data from the current row
    for paragraph in doc.paragraphs:
        replace_placeholders(paragraph, row)

    # Save the merged document in the output directory
    output_file = os.path.join(output_dir, f'merged_document_{index+1}.docx')
    doc.save(output_file)

print('Mail merge completed.')