# Multi-File CSV Converter

## The script will:

- Process all PDF, DOCX, and PPTX files in their respective directories.
- Save each file's extracted content into a CSV file.
- Log any errors encountered during the conversion process.

## Modularization

The script is modularized, allowing for easy extension.

- Add more file types by creating new processing functions.
- The process_directory function generalizes the directory traversal and file handling.

## Error Logging

- All errors during processing are logged in the error_log.txt file.
- Each log entry includes the file name and a detailed error message.

## Structure

### Main Functions

- **`process_pdf_file`**:

  - Converts PDF pages to images using `pdf2image`.
  - Extracts text from each image using Tesseract OCR.
  - Stores extracted text in a list of lists (each list represents a line of text).

- **`process_docx_file`**:

  - Reads DOCX files using the `python-docx` library.
  - Extracts paragraphs from the document and stores them in a list of lists.

- **`process_ppt_file`**:
  - Reads PPT files using `python-pptx`.
  - Extracts text from slides and stores it in a list of lists.

### Helper Functions

- **`write_to_csv`**:

  - Writes the extracted data (from PDF, DOCX, or PPT files) to a CSV file.

- **`log_error`**:
  - Logs any errors encountered during file processing to an `error_log.txt` file.

### Process Directory

- **`process_directory`**:
  - Loops through the files in a specified directory.
  - Processes files based on their type (PDF, DOCX, PPTX) by calling the respective function (`process_pdf_file`, `process_docx_file`, `process_ppt_file`).

### Main Execution

- The `main` function coordinates the overall execution by:
  1. Clearing any existing `error_log.txt` file.
  2. Calling `process_directory` for each file type (PDF, DOCX, PPTX).
  3. Processing all files and writing the extracted data into CSV files.

## Error Handling and Logging

- If any file encounters an error during processing, it will be logged to `error_log.txt`.
- Each error entry contains the file name and a description of the issue.

## Directory Structure Assumption

- **PDF Files**: Located in `Mezmure/PDF`.
- **DOCX Files**: Located in `Mezmure/WORD`.
- **PPT Files**: Located in `Mezmure/PPT`.
- **CSV Output**: Files will be saved in `Mezmure/CSV_Output`.

### CSV_Output

Each page of a file is now formated as a single row in the CSV file.

- Columns are:
- **Mezmur Name**: (from the file name),
- **Page/Verse**: (page number),
- **Zemari Name**: (default to "Unknown Zemari" unless specified),
- **File Name**: (full path to the file),
- **Verses**: (extracted text).

## Requirements

Ensure that the following libraries are installed:

- `pytesseract`
- `pdf2image`
- `python-docx`
- `python-pptx`

Additionally, make sure that:

- **Tesseract OCR** is installed and configured with Amharic language support (`amh.traineddata`).

### Install Required Libraries

```bash
pip install pytesseract pdf2image python-docx python-pptx
```

Install Tesseract OCR
On macOS, you can install Tesseract using Homebrew:

```bash
brew install tesseract-lang
```

Running the Script
Once all dependencies are installed, you can run the script:

```bash
python file_converter.py
```

```bash

### How to Use:
1. Update the script's file directories if needed.
2. Follow the steps for installing dependencies and Tesseract OCR.
3. Run the Python script as described.

Let me know if you'd like any further changes!

```
