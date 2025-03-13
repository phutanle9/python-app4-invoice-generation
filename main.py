import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Ensure that the PDFs directory exists, if not, create it
output_dir = Path("PDFs")
output_dir.mkdir(parents=True, exist_ok=True)  # Create the directory if it doesn't exist

# Get all file paths with .xlsx extension in the invoices folder
filepaths = glob.glob("invoices/*.xlsx")

# Loop through each file path and read the corresponding Excel file
for filepath in filepaths:
    try:
        # Read the sheet named "Sheet 1"
        df = pd.read_excel(filepath, sheet_name="Sheet 1")  # Read the sheet named "Sheet 1"

        # Print the DataFrame for debugging purposes
        print(df)  # Print the DataFrame to check the content

        # Initialize PDF
        pdf = FPDF(orientation="P", unit="mm", format="A4")
        pdf.add_page()

        # Extract filename and invoice number from the filepath
        filename = Path(filepath).stem  # Get the filename without the extension
        invoice_nr = filename.split("-")[0]  # Assuming the invoice number is the first part of the filename

        # Set the font for the invoice number
        pdf.set_font(family="Times", size=16, style="B")

        # Add the invoice number to the PDF
        pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr}")

        # Save the PDF to the output directory
        pdf.output(output_dir / f"{filename}.pdf")  # Save PDF in the PDFs folder

    except Exception as e:
        print(f"Error processing file {filepath}: {e}")
