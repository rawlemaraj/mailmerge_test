import pandas as pd
from openpyxl import load_workbook
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os

# Load Excel data
filename = 'your_excel_file.xlsm'  # replace with your file
wb = load_workbook(filename, data_only=True)  # data_only=True to evaluate formulae
ws = wb.active  # get active sheet
data = pd.DataFrame(ws.values)  # convert to pandas DataFrame

# Assume that the first row is the header
header = data.iloc[0]
data = data[1:]
data.columns = header

# Load your Word template
template = Document('your_word_template.docx')  # replace with your file

# Now let's go through each row and replace placeholders with values from Excel
for index, row in data.iterrows():
    doc = template.clone()  # clone the template to make a new document
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            text = run.text
            for key, value in row.iteritems():
                if "{" + str(key) + "}" in text:
                    text = text.replace("{" + str(key) + "}", str(value))
                    run.text = text

    # Save the document
    doc.save(os.path.join('output', f'document_{index}.docx'))
