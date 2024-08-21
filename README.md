# PDF Table Extraction and Excel Export Tool

Extract tables from a specified PDF file and saves them into an Excel file, with each table placed on a separate sheet. The script leverages `pdfplumber` for PDF processing and `pandas` for data handling and exporting.

## Prerequisites

Ensure you have the following Python libraries installed:

- `pdfplumber`
- `pandas`
- `xlsxwriter`

You can install these dependencies using pip:
pip install pdfplumber pandas xlsxwriter

## Usage
Set the PDF File Path:
Modify the pdf_file variable in the script to the absolute path of the PDF file you want to extract tables from.
pdf_file = r"path\to\your\file.pdf"


## Run the Script:
Execute the script to:

## Open the specified PDF file.
Extract tables from the second page of the PDF.
Save each table to a separate sheet in an Excel file named Extracted.xlsx located in the same directory as the PDF.
Print each table's content to the console.
Output:

An Excel file named Extracted.xlsx will be created in the same directory as the PDF file. Each sheet will be named table_1, table_2, etc., corresponding to the extracted tables.
The script will also print the DataFrames representing each table to the console.
Customization
Page Selection: The script extracts tables from the second page (pdf.pages[1]). Modify the index number to select a different page.

DataFrame Cleanup: The script includes a step to remove extra spaces from string entries in the tables. Adjust or remove this based on the structure of your tables.

Excel File Path: The output Excel file is saved in the same directory as the PDF file. You can change the excel_file variable to save it elsewhere.

## Error Handling
The script includes basic error handling to catch and display any issues that arise during PDF processing.

## Example
For a PDF containing 3 tables on the second page, the script will generate an Excel file with 3 sheets: table_1, table_2, and table_3.

