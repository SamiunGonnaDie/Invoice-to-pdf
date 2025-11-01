import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path


filepaths = glob.glob('invoices/*.xlsx')

for filepath in filepaths:


    pdf = FPDF("p", "mm", "A4")
    pdf.add_page()


    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")


    pdf.set_font("Times", size=12, style="B")
    pdf.cell(w= 50, h=8, text=f"Invoice nr.{invoice_nr}", ln=1)

    pdf.set_font("Times", size=12, style="B")
    pdf.cell(w= 50, h=8, text=f"Date:{date}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    # Add a header
    columns = list(df.columns)
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font("Times", size=12, style="B")
    pdf.set_draw_color(80, 80, 80)
    pdf.cell(w=30, h=8, text= columns[0], border=1)
    pdf.cell(w=50, h=8, text= columns[1], border=1)
    pdf.cell(w=50, h=8, text= columns[2], border=1)
    pdf.cell(w=30, h=8, text= columns[3], border=1)
    pdf.cell(w=30, h=8, text= columns[4], border=1, ln=1)

    #Add rows to the table
    for index, row in df.iterrows():
        pdf.set_font("Times", size=12)
        pdf.set_draw_color(80,80,80)
        pdf.cell(w= 30, h=8, text=str(row["product_id"]), border=1)
        pdf.cell(w= 50, h=8, text=str(row["product_name"]), border=1)
        pdf.cell(w= 50, h=8, text=str(row["amount_purchased"]), border=1)
        pdf.cell(w= 30, h=8, text=str(row["price_per_unit"]), border=1)
        pdf.cell(w= 30, h=8, text=str(row["total_price"]), border=1, ln=1)

    total_sum = df["total_price"].sum()
    pdf.set_font("Times", size=12)
    pdf.set_draw_color(80, 80, 80)
    pdf.cell(w=30, h=8, text="", border=1)
    pdf.cell(w=50, h=8, text="", border=1)
    pdf.cell(w=50, h=8, text="", border=1)
    pdf.cell(w=30, h=8, text="", border=1)
    pdf.cell(w=30, h=8, text= str(total_sum), border=1, ln=1)

    # Add text outside box
    pdf.set_font("Times", size=12, style="B")
    pdf.cell(w= 50, h=8, text=f"The total price is {total_sum}", ln=1)

    # name & logo
    pdf.set_font("Times", size=12,style="B")
    pdf.cell(w= 25, h=8, text=f"PythonHow")
    pdf.image("004 pythonhow.png", w=10, h=10)


    pdf.output(f"PDFs/{filename}.pdf") 








