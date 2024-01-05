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
    pdf.cell(w=50, h=8, txt=f"Dated : {date}",ln=2)

    # Add Header for the table
    columns = df.columns
    # In order to remove the Underscores in the Header row of the table.
    columns = [item.replace("_"," ").title() for item in columns]
    # To set the color an font for the title
    pdf.set_font(family="Times",size=12,style="B")
    pdf.set_text_color(60,60,60)
    pdf.cell(w=30,h=10,txt=columns[0],border=1)
    pdf.cell(w=50,h=10,txt=columns[1],border=1)
    pdf.cell(w=50,h=10,txt=columns[2],border=1)
    pdf.cell(w=30,h=10,txt=columns[3],border=1)
    pdf.cell(w=30,h=10,txt=columns[4],border=1,ln=1)

    # Add rows
    for index,row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(60, 60, 60)
        pdf.cell(w=30,h=10,txt=str(row["product_id"]),border=1)
        pdf.cell(w=50,h=10,txt=str(row["product_name"]),border=1)
        pdf.cell(w=50,h=10,txt=str(row["amount_purchased"]),border=1)
        pdf.cell(w=30,h=10,txt=str(row["price_per_unit"]),border=1)
        pdf.cell(w=30,h=10,txt=str(row["total_price"]),border=1,ln=1)



    pdf.output(f"PDFs/{filename}.pdf")






