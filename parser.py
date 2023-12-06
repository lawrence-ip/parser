import pandas as pd
from docx import Document

# Read the Excel file
excel_file = 'path/to/excel/file.xlsx'
df = pd.read_excel(excel_file)

# Extract the numbers from the Excel data
numbers = df['Numbers'].tolist()

# Create a Word document
doc = Document()

# Write the numbers to the Word document
for number in numbers:
    doc.add_paragraph(str(number))

# Save the Word document
word_file = 'path/to/output/document.docx'
doc.save(word_file)

# Alternatively, save the numbers to a Markdown file
markdown_file = 'path/to/output/document.md'
with open(markdown_file, 'w') as file:
    for number in numbers:
        file.write(str(number) + '\n')