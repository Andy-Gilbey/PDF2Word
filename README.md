# PDF to DOCX Converter

## Description
This Python script converts PDF documents to DOCX format by extracting text from the PDF and saving it as a Word document. 
It works on the command-line with user prompts to specify the PDF file path and then handles the conversion process.

## Required Libraries
To run this script, the following Python libraries are required:
- `io`: Python library for i/o operations.
- `os`: Library which interacts with the operating system.
- `re`: Library which works utilising expression operations.
- `pdfminer`: Library for extracting text from PDF files.
- `python-docx`: Library for creating/updating Word documents.

You can install the necessary libraries using pip:

```bash
pip install pdfminer.six python-docx
```

## Script Functions
### printBanner()
Displays a banner at the start of the program.

### extractTextFromPdf(pdfPath)
Extracts all text from the specified PDF file. It works on a  page-by-page basis, processing each in sequence and displays the progress of the extraction.

### removeNonXmlCharacters(text)
Cleans the taken text by removing any characters that are not compatible with XML (used in DOCX files).

### getGoodbyeMessage()
Randomly selects a goodbye message from a predefined list with varied language options (weighted towards English).
Just a fun addition.

## Running the Script
To use the script, follow these steps:
1. Ensure all required libraries are installed.
2. Save the script to a file, for example, `pdf_to_docx.py`.
3. Open a terminal or command prompt.
4. Navigate to the directory containing the script.
5. Run the script by executing:

```bash
python pdf_to_docx.py
```

Follow the prompts to input the full path of the PDF file you wish to convert or type 'q' to quit the program. 
You can get the full path of a file on a linux based system using the command 

```bash
realpath
```

The script will then process the PDF, convert it to text, and save the text in a DOCX file located in the directory which the orinigal PDF lives.

