from mailmerge import MailMerge
import sqlite3
from pathlib import Path

# Using docx2pdf
from docx2pdf import convert

# Alternative: If using Aspose
'''
import aspose.words as aw
'''

# Alternative: If using LibreOffice
'''
import subprocess
'''

# Connect to your SQLite database
conn = sqlite3.connect('employees.sqlite')
cursor = conn.cursor()

# Fetch data from the SQL table
cursor.execute("SELECT * FROM employees")
data = cursor.fetchall()

# Create output directories
output_docx_dir = Path.cwd() / 'output' / 'docx'
output_docx_dir.mkdir(parents=True, exist_ok=True)
output_pdf_dir = Path.cwd() / 'output' / 'pdf'
output_pdf_dir.mkdir(parents=True, exist_ok=True)

# Loop through each row of data
for i, row in enumerate(data):
    # Create a Word document from the template
    doc = MailMerge('template.docx')

    # Get column names
    columns = [description[0] for description in cursor.description]

    # Create a dictionary mapping column names to values
    row_dict = dict(zip(columns, row))
    for key in row_dict:
        row_dict[key] = str(row_dict[key])

    print(row_dict)

    doc.merge_pages([row_dict])

    # Save the populated Word document
    doc_output = output_docx_dir / f'temp_document_{i}.docx'
    doc.write(doc_output)

    pdf_output = output_pdf_dir / f'output_{i}.pdf'

    ## Use docx2pdf, requires manual attention on mac. Might work on Windows for free.
    convert(doc_output, pdf_output)

    ## Alternative: Use Aspose, requires 1200$ license to get rid of watermarks.
    '''
    # doc2 = aw.Document(doc_output)
    # doc2.save(pdf_output)
    '''

    ## Alternative: Use LibreOffice with good result and free, requires installation.
    '''
    rc = subprocess.call(['which', 'soffice'])
    if rc != 0:
        print("Libreoffice is not installed. Use 'brew install --cask libreoffice'.")
        exit(1)

    output = subprocess.check_output(['soffice', '--convert-to', 'pdf', doc_output, '--outdir', output_pdf_dir])
    print(output)
    '''

# Close the database connection
conn.close()
