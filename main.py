import pandas as pd
import glob
from fpdf import FPDF

# Read excel files
filepaths = glob.glob("invoices/*.xlsx")

# Write each excel in PDF file
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{filepath[9:14]}", align="L", ln=1, border=0)
    pdf.cell(w=50, h=8, txt=f"Date {filepath[15:24]}", align="L", ln=1, border=0)
    pdf.output(f"PDF's/{filepath[9:24]}.pdf")