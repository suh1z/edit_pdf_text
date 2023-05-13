import os
from pdf2docx import Converter
from docx import Document
from docx2pdf import convert

def convert_pdf_to_docx(input_pdf_path, output_docx_path):
    cv = Converter(input_pdf_path)
    cv.convert(output_docx_path, start=0, end=None)
    cv.close()

def replace_word_in_docx(docx_path, target_word, replacement_word):
    doc = Document(docx_path)
    for paragraph in doc.paragraphs:
        if target_word in paragraph.text:
            for run in paragraph.runs:
                run.text = run.text.replace(target_word, replacement_word)
    doc.save("modified.docx")

def convert_docx_to_pdf(input_docx_path, output_pdf_path):
    convert(input_docx_path, output_pdf_path)

# Get user input for the target word and replacement word
target_word = input("Enter the word to replace: ")
replacement_word = input("Enter the replacement word: ")

# Define input and output file paths
input_pdf_path = 'input.pdf'
input_dir = os.path.dirname(input_pdf_path)
output_docx_path = os.path.join(input_dir, 'output.docx')
modified_docx_path = os.path.join(input_dir, 'modified.docx')
output_pdf_path = os.path.join(input_dir, 'final_output.pdf')

# Convert PDF to DOCX
convert_pdf_to_docx(input_pdf_path, output_docx_path)

# Replace word in DOCX
replace_word_in_docx(output_docx_path, target_word, replacement_word)

# Convert modified DOCX to PDF
convert_docx_to_pdf(modified_docx_path, output_pdf_path)

# Remove intermediate files
os.remove(output_docx_path)
os.remove(modified_docx_path)

print("Conversion completed successfully!")
