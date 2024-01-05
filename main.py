import pandas as pd
import glob
from fpdf import FPDF

# Use to retrieve the data from excel into a variable
filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath,sheet_name="Sheet 1")
    pdf = FPDF(orientation="p", unit="mm", format="A4")
    pdf.add_page()

    print(df)






