import os
import pdfplumber
import pandas as pd

# Get the absolute path to the PDF file
pdf_file = r"ast_sci_data_tables_sample.pdf"

# Get the directory of the PDF file
pdf_dir = os.path.dirname(os.path.abspath(pdf_file))

try:
    # Open the PDF file with pdfplumber
    with pdfplumber.open(pdf_file) as pdf:
        # Extract the first page
        page = pdf.pages[1]

        # Extract tables from the page
        tables = page.extract_tables()
        print("Tables extracted:", len(tables))
except Exception as e:
    print(f"Error: {e}")

else:
    # Create an Excel writer object
    excel_file = os.path.join(pdf_dir, "Extracted.xlsx")
    with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
        # Save each table to a separate sheet in the Excel file
        for i, table in enumerate(tables, start=1):
            df = pd.DataFrame(table)
            sheet_name = f"table_{i}"

            # Clean up DataFrame (optional: depending on table formatting)
            df = df.applymap(lambda x: ' '.join(str(x).split()) if isinstance(x, str) else x)

            # Write to Excel sheet
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"Excel sheet generated: {sheet_name}")

            # Print the DataFrame
            print(f"DataFrame for {sheet_name}:\n", df)

    print(f"Excel file saved at: {excel_file}")
