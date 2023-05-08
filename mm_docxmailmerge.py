import openpyxl
from mailmerge import MailMerge

# Load the Excel data
def load_excel_data(file_path):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    headers = [cell.value for cell in sheet[1]]
    data = [dict(zip(headers, [cell.value for cell in row])) for row in sheet.iter_rows(min_row=2)]

    return data

# Perform mail merge
def mail_merge(template_path, data, output_folder):
    for index, record in enumerate(data, start=1):
        with MailMerge(template_path) as document:
            document.merge(**record)
            output_file = f"{output_folder}/merged_document_{index}.docx"
            document.write(output_file)

# Configuration
excel_file = 'data.xlsx'
word_template = 'mail_merge_template.dotx'
output_folder = 'output'

# Load data from Excel and perform mail merge
data = load_excel_data(excel_file)
mail_merge(word_template, data, output_folder)