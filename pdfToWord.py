import random
import io  # Input/Output library
import os  # Operating system library
import re  # Regular expression library
from pdfminer.converter import TextConverter  # PDF to text converter
from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager  # PDF interpreter and resource manager
from pdfminer.layout import LAParams  # Layout parameters for PDF conversion
from pdfminer.pdfpage import PDFPage  # PDF page handling
from docx import Document  # Word document library

# Function to print the program banner
def printBanner():
    banner = """
   ▄███████▄ ████████▄     ▄████████          ███     ███    █▄        ▄█     █▄   ▄██████▄     ▄████████ ████████▄  
  ███    ███ ███   ▀███   ███    ███      ▀█████████▄ ███    ███      ███     ███ ███    ███   ███    ███ ███   ▀███ 
  ███    ███ ███    ███   ███    █▀          ▀███▀▀██ ███    ███      ███     ███ ███    ███   ███    ███ ███    ███ 
  ███    ███ ███    ███  ▄███▄▄▄              ███   ▀ ███    ███      ███     ███ ███    ███  ▄███▄▄▄▄██▀ ███    ███ 
▀█████████▀  ███    ███ ▀▀███▀▀▀              ███     ███    ███      ███     ███ ███    ███ ▀▀███▀▀▀▀▀   ███    ███ 
  ███        ███    ███   ███                 ███     ███    ███      ███     ███ ███    ███ ▀███████████ ███    ███ 
  ███        ███   ▄███   ███                 ███     ███    ███      ███     ███ ███    ███   ███    ███ ███   ▄███ 
 ▄████▀      ████████▀    ███                ▄████▀   ████████▀        ▀███▀███▀   ▀██████▀    ███    ███ ████████▀                                                                                                                                                                                                                                                                                                                                                                    
    """
    print(banner)


def extractTextFromPdf(pdfPath):
    resourceManager = PDFResourceManager()  # Create a PDF resource manager
    fakeFileHandle = io.StringIO()  # Create a fake file handle
    converter = TextConverter(resourceManager, fakeFileHandle, laparams=LAParams())  # Create a text converter
    pageInterpreter = PDFPageInterpreter(resourceManager, converter)  # Create a page interpreter

    totalPages = 0  # Total number of pages in the PDF
    processedPages = 0  # Number of pages processed

    with open(pdfPath, 'rb') as fileHandle:
        totalPages = sum(1 for _ in PDFPage.get_pages(fileHandle))  # Calculate total pages in the PDF

        fileHandle.seek(0)  # Reset the file handle position for processing

        for page in PDFPage.get_pages(fileHandle, caching=True, check_extractable=True):
            # Process each page in the PDF
            pageInterpreter.process_page(page)
            processedPages += 1

            progress = processedPages / totalPages * 100  # Calculate progress as a percentage
            print(f"Progress: {progress:.2f}%")  # Display progress

        text = fakeFileHandle.getvalue()  # Get the extracted text from the fake file handle

    converter.close()  # Close the converter
    fakeFileHandle.close()  # Close the fake file handle

    if text:
        return text


def removeNonXmlCharacters(text):
    # Remove non-XML compatible characters using a regular expression
    xml_compatible_text = re.sub(r'[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD]+', '', text)
    return xml_compatible_text


def getGoodbyeMessage():
    messages = {
        'en': 'Goodbye!',
        'es': '¡Adiós!',
        'fr': 'Au revoir!',
        'ja': 'さようなら！'
    }
    language = random.choices(list(messages.keys()), weights=[0.45, 0.15, 0.15, 0.25], k=1)[0] # Gotta ensure that the english goodbye has more probability of appearing then the others
    return messages[language]


if __name__ == '__main__':
    try:
        printBanner()

        while True:
            pdfPath = input("Enter the FULL path to the PDF file (q to quit): ").strip('\"')  # Prompt the user for the PDF file path and remove double quotation marks
            # Check if the user wants to quit
            if pdfPath.lower() == 'q':
                print(getGoodbyeMessage())
                break

            # Check if the file exists
            if not os.path.isfile(pdfPath):
                print("File not found!")
                continue

            wordFilename = os.path.splitext(os.path.basename(pdfPath))[0] + '_CONVERTED.docx'
            wordPath = os.path.join(os.path.dirname(pdfPath), wordFilename)  # Create the output file path

            extractedText = extractTextFromPdf(pdfPath)  # Extract text from the PDF

            if not extractedText:
                print("No text extracted from the PDF!")
                continue

            extractedText = removeNonXmlCharacters(extractedText)  # Remove non-XML compatible characters

            doc = Document()  # Create a new Word document
            doc.add_paragraph(extractedText)  # Add the extracted text to the document as a paragraph
            doc.save(wordPath)  # Save the Word document

            print("Conversion completed successfully!")  # Display success message

    except ValueError as e:
        print(f"Error: {str(e)}")  # value error and display error message
    except FileNotFoundError as e:
        print(f"Error: {str(e)}")  #  file not found error and display error message
    except KeyboardInterrupt:
        print("\nProgram interrupted by the user!") # ctrl + c handle 
        print(getGoodbyeMessage())
    except Exception as e:
        print(f"An error occurred: {str(e)}")  #  any other exception and display error message
