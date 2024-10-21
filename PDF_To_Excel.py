"""
Python 3.12.6

User Input:
    The PDF page number to read after executing the program.

Description:
    Extracts data from specified PDF pages and writes it to Excel.
    Each bibliographic entry consists of a Chinese title and an English title,
    arranged in separate columns of a single row.

Developer:
    Shiuan-Jen, Yang
"""
import pdfplumber
import pandas as pd
import re
import os

# Extract table of contents titles from a specified PDF page
def extract_titles_from_toc(pdf_path, page_number):
    extracted_data = []
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[page_number - 1]
        text = page.extract_text()
        if text:
            lines = text.split('\n')
            i = 0
            while i < len(lines):
                line = lines[i].strip()
                english_title = ""

                # Chinese Title: Check for at least two consecutive ellipses ("...") or ellipsis ("…"), indicating the end of the title;
                # otherwise, continue searching for the remaining title content.
                if any('\u4e00' <= char <= '\u9fff' for char in line) and re.search(r'(\.\.\.|…){2,}', line):
                    chinese_title = re.split(r'(\.\.\.|…){2,}', line)[0].strip()
                    chinese_title = re.sub(r'\s*\d+$', '', chinese_title).strip()
                    i += 1
                    continue

                # English Title: Check for at least spaces, indicating the end of the title;
                # otherwise, continue searching for the remaining title content.
                if line.isascii():
                    english_title += line + " "
                    if line.endswith("  "):
                        extracted_data.append([chinese_title, english_title.strip()])
                        i += 1
                    else:
                        i += 1
                    continue

                # Check if the line is a page number, if so, stop processing
                if re.match(r'^\s*(\d+|[ivxlcdm]+)\s*$', line.lower()):
                    break
                i += 1

    return extracted_data

# Save the extracted data to an Excel file
def save_to_excel(data, excel_path):
    df = pd.DataFrame(data, columns=["Title", "Alternative Title (1)"])
    df.index += 2
    df.to_excel(excel_path, index=False)

def main():
   
    data = extract_titles_from_toc(pdf_path, page_number)
    if data:
        save_to_excel(data, excel_path)
        print(f"Successfully saved titles from page {page_number} to {excel_path}")
    else:
        print(f"No titles extracted from page {page_number}")

    # Attempt to open the Excel file
    try:
        os.startfile(excel_path)
    except Exception as e:
        print(f"Unable to open the file automatically: {e}")

# Main program
if __name__ == '__main__':
    pdf_path = 'Bibliography to be Created.pdf'
    excel_path = 'Bibliography_List_Example.xlsx'
    page_number = int(input("Enter the page number to read: "))

    main()

