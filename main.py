import pandas as pd
import glob
from fpdf import FPDF

# Read excel files
filepaths = glob.glob("invoices/*.xlsx")

# Write each excel in PDF file
for filepath in filepaths:
    # Add PDF file for each excel file
    pdf = FPDF(orientation="P", unit="mm", format="A4")

    # Extract info from excel files
    pdf.add_page()
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(
        w=50, h=8, txt=f"Invoice nr.{filepath[9:14]}", align="L", ln=1, border=0)
    pdf.cell(
        w=50, h=8, txt=f"Date {filepath[15:24]}", align="L", ln=1, border=0)

    # Read from excel
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    columns_name = df.columns
    columns_name = [item.replace("_", " ").capitalize()
                    for item in columns_name]

    # Agregar nombre de columnas
    pdf.set_font(family="Times", style="B", size=10)
    pdf.cell(w=30, h=8, txt=str(columns_name[0]), ln=0, border=1)
    pdf.cell(w=60, h=8, txt=str(columns_name[1]), ln=0, border=1)
    pdf.cell(w=40, h=8, txt=str(columns_name[2]), ln=0, border=1)
    pdf.cell(w=30, h=8, txt=str(columns_name[3]), ln=0, border=1)
    pdf.cell(w=30, h=8, txt=str(columns_name[4]), ln=1, border=1)

    # Agregar filas
    for idx, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), ln=0, border=1)
        pdf.cell(w=60, h=8, txt=str(row["product_name"]), ln=0, border=1)
        pdf.cell(w=40, h=8, txt=str(row["amount_purchased"]), ln=0, border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), ln=0, border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), ln=1, border=1)

    # Agregar ultima fila
    total_price = df["total_price"].sum()
    pdf.set_font(family="Times", style="B", size=10)
    pdf.cell(w=30, h=8, txt="", ln=0, border=1)
    pdf.cell(w=60, h=8, txt="", ln=0, border=1)
    pdf.cell(w=40, h=8, txt="", ln=0, border=1)
    pdf.cell(w=30, h=8, txt="", ln=0, border=1)
    pdf.cell(w=30, h=8, txt=str(total_price), ln=1, border=1)

    # Agregar texto final
    pdf.set_font(family="Times", style="B", size=14)
    pdf.cell(
        w=30, h=8, txt=f"The total price is: {total_price}", ln=1, border=0)
    pdf.cell(w=30, h=8, txt=f"PythonHow", ln=0, border=0)
    pdf.image("pythonhow.png", w=10)

    # Create PDF files
    pdf.output(f"PDF's/{filepath[9:24]}.pdf")
