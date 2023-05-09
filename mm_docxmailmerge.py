# pip install python-docx-mailmerge openpyxl

import openpyxl
from mailmerge import MailMerge

# Load the Excel data
def load_excel_data(file_path):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    headers = [cell.value for cell in sheet[1]]
    data = [dict(zip(headers, [cell.value for cell in row])) for row in sheet.iter_rows(min_row=2)]

    return data

# Get a single record from the data based on a unique value
def get_record_by_unique_value(data, unique_value, unique_column):
    for record in data:
        if record[unique_column] == unique_value:
            return record
    return None

# Perform mail merge
def mail_merge(template_path, record, output_file):
    with MailMerge(template_path) as document:
        document.merge(**record)
        document.write(output_file)

# Configuration
excel_file = 'data.xlsx' # define patch to file 
word_template = 'mail_merge_template.dotx' # define path to file

# Load data from Excel
data = load_excel_data(excel_file)

# Prompt the user to enter the unique value and column name
unique_value = input('Please enter the unique value: ')
unique_column = input('Please enter the column name where the unique value is located: ')

# Prompt the user to enter the output file name
output_file_name = input('Please enter the output file name: ')
output_file = f'{output_file_name}.docx'

record = get_record_by_unique_value(data, unique_value, unique_column)

if record is not None:
    # Perform mail merge with the chosen row
    mail_merge(word_template, record, output_file)
    print(f'Mail merge completed for {unique_column}: {unique_value}')
else:
    print(f'No record found with {unique_column}: {unique_value}')


# A new function get_record_by_id is added to retrieve a single record from the data based on a unique ID value and the column name where the ID is located.
# The configuration variables are updated to remove the output_file variable.
# The input function is used to prompt the user to enter the unique value, column name, and output file name.
# The output file name is created by appending the ".docx" file extension to the user-provided name.
# The mail merge is performed using the chosen row.