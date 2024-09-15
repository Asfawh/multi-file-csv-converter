import pytesseract
from pdf2image import convert_from_path
from docx import Document
import os
import csv
import mimetypes
from pptx import Presentation

# Directories
pdf_directory = 'Mezmure/PDF'
docx_directory = 'Mezmure/WORD'
ppt_directory = 'Mezmure/PPT'
output_directory = 'Mezmure/CSV_Output1'
error_log_file = os.path.join(output_directory, 'error_log.txt')

# Ensure that the output directory exists
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

# Function to log errors
def log_error(file_name, error_message):
    with open(error_log_file, 'a') as error_log:
        error_log.write(f"Error processing {file_name}: {error_message}\n")

# Function to process PDFs
def process_pdf_file(pdf_file_path):
    try:
        pages = convert_from_path(pdf_file_path)
        data = []
        for page in pages:
            text = pytesseract.image_to_string(page, lang='amh')
            lines = text.split('\n')
            for line in lines:
                data.append([line])

        return data
    except Exception as e:
        raise RuntimeError(f"Error processing PDF: {str(e)}")

# Function to process DOCX files
def process_docx_file(docx_file_path):
    try:
        doc = Document(docx_file_path)
        data = []
        for paragraph in doc.paragraphs:
            data.append([paragraph.text])
        return data
    except Exception as e:
        raise RuntimeError(f"Error processing DOCX: {str(e)}")

# Function to process PPT files
def process_ppt_file(ppt_file_path):
    try:
        prs = Presentation(ppt_file_path)
        data = []
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    data.append([shape.text])
        return data
    except Exception as e:
        raise RuntimeError(f"Error processing PPT: {str(e)}")

# Function to write to CSV
def write_to_csv(data, csv_file_path):
    try:
        with open(csv_file_path, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerows(data)
    except Exception as e:
        raise RuntimeError(f"Error writing CSV: {str(e)}")

# Function to process all files in a directory
def process_directory(directory, file_type, process_function):
    for file_name in os.listdir(directory):
        file_path = os.path.join(directory, file_name)
        if file_name.endswith(file_type):
            try:
                # Process the file
                extracted_data = process_function(file_path)
                # Create CSV file name
                csv_file_name = os.path.splitext(file_name)[0] + '.csv'
                csv_file_path = os.path.join(output_directory, csv_file_name)
                # Write extracted data to CSV
                write_to_csv(extracted_data, csv_file_path)
                print(f"Successfully processed and converted {file_name} to {csv_file_name}")
            except Exception as e:
                log_error(file_name, str(e))
                print(f"Error logged for {file_name}")

# Main function to run all the processing
def main():
    # Clear previous error log
    if os.path.exists(error_log_file):
        os.remove(error_log_file)

    # Process PDF files
    print("Processing PDF files...")
    process_directory(pdf_directory, '.pdf', process_pdf_file)

    # Process DOCX files
    print("Processing DOCX files...")
    process_directory(docx_directory, '.docx', process_docx_file)

    # Process PPT files
    print("Processing PPT files...")
    process_directory(ppt_directory, '.pptx', process_ppt_file)

    print("All files have been processed.")

if __name__ == "__main__":
    main()
