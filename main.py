import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filenames = glob.glob("invoices/*.xlsx")
pdf = FPDF(orientation="P", unit="mm", format="A4")

for filename in filenames:
    data = pd.read_excel(filename, sheet_name = "Sheet 1")
    columns = list(data.columns)
    columns = [item.replace("_"," ") for item in columns]
    pdf.add_page()

    filename = Path(filename).stem
    invoice_nr, date = filename.split("-")


    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=30, h=12, txt="Invoice nr." + invoice_nr, align="L", ln=1)

    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=30, h=12, txt="Date :" + date, align="L", ln=1)


    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=12, txt=columns[0], align="L", border=1)
    pdf.cell(w=60, h=12, txt=columns[1], align="L", border=1)
    pdf.cell(w=35, h=12, txt=columns[2], align="L", border=1)
    pdf.cell(w=30, h=12, txt=columns[3], align="L", border=1)
    pdf.cell(w=30, h=12, txt=columns[4], align="L", border=1, ln=1)

    for index, row in data.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.cell(w=30, h=12, txt=str(row["product_id"]), align="L",border=1)
        pdf.cell(w=60, h=12, txt=str(row["product_name"]), align="L", border=1)
        pdf.cell(w=35, h=12, txt=str(row["amount_purchased"]), align="L", border=1)
        pdf.cell(w=30, h=12, txt=str(row["price_per_unit"]), align="L", border=1)
        pdf.cell(w=30, h=12, txt=str(row["total_price"]), align="L", border=1, ln=1)

    pdf.output(f"PDFs/{filename}.pdf")