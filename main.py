import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Use to retrieve the data from excel into a variable
filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    # Reads the Excel sheet
    df = pd.read_excel(filepath,sheet_name="Sheet 1")
    pdf = FPDF(orientation="p", unit="mm", format="A4")
    # Add pdf page
    pdf.add_page()

    # This helps us to filename from the directory
    filename = Path(filepath).stem
    # Helps to get the invoice number and date by split function
    invoice_no,date = filename.split("-")
    # This is to display the Invoice Number
    pdf.set_font(family="Times",size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice Number: {invoice_no}",ln= 1)

    # This is to display the date
    pdf.set_font(family="Times",size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Dated : {date}")


    pdf.output(f"PDFs/{filename}.pdf")






