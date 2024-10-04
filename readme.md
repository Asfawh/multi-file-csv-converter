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
  `~/Documents/repo/Py/multi-file-csv-converter main !9 ?1 ❯ pipenv shell` #Py-kyBhG0Ms

2. Follow the steps for installing dependencies and Tesseract OCR.
  `❯ pip list
  Package            Version
  ------------------ -----------
  aiohttp            3.9.5
  aiohttp-retry      2.8.3
  aiosignal          1.3.1
  appnope            0.1.4
  asttokens          2.4.1
  async-timeout      4.0.3
  attrs              23.2.0
  certifi            2024.2.2
  cffi               1.17.1
  charset-normalizer 3.3.2
  comm               0.2.2
  cryptography       43.0.1
  debugpy            1.8.1
  decorator          5.1.1
  distro             1.9.0
  exceptiongroup     1.2.1
  executing          2.0.1
  frozenlist         1.4.1
  idna               3.7
  ipykernel          6.29.4
  ipython            8.24.0
  jedi               0.19.1
  jupyter_client     8.6.1
  jupyter_core       5.7.2
  lxml               5.3.0
  matplotlib-inline  0.1.7
  multidict          6.0.5
  nest-asyncio       1.6.0
  numpy              2.1.1
  packaging          24.0
  pandas             2.2.2
  parso              0.8.4
  pdf2image          1.17.0
  pdfminer.six       20231228
  pdfplumber         0.11.4
  pexpect            4.9.0
  pillow             10.4.0
  pip                24.0
  platformdirs       4.2.2
  prompt-toolkit     3.0.43
  psutil             5.9.8
  ptyprocess         0.7.0
  pure-eval          0.2.2
  pycparser          2.22
  Pygments           2.18.0
  PyJWT              2.8.0
  pypandoc           1.13
  PyPDF2             3.0.1
  pypdfium2          4.30.0
  pytesseract        0.3.13
  python-dateutil    2.9.0.post0
  python-docx        1.1.2
  python-pptx        1.0.2
  pytz               2024.2
  pyzmq              26.0.3
  requests           2.31.0
  setuptools         69.2.0
  six                1.16.0
  stack-data         0.6.3
  tabula-py          2.9.3
  tornado            6.4
  traitlets          5.14.3
  twilio             9.0.5
  typing_extensions  4.11.0
  tzdata             2024.1
  urllib3            2.2.1
  wcwidth            0.2.13
  wheel              0.43.0
  XlsxWriter         3.2.0
  yarl               1.9.4`

3. Run the Python script as described.
  `~/Documents/repo/Py/multi-file-csv-converter main !9 ?2 ❯ python file_converter.py`

```
