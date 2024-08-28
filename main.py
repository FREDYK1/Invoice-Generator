import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob('Invoices/*.xlsx')

for filepath in filepaths:
    if Path(filepath).name.startswith('~$'):
        continue

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()


    filename = Path(filepath).stem
    InvoiceName, Date = filename.split("-")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=10, txt=f"Invoice nr: {InvoiceName}", border=0, ln=1, align="L")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=10, txt=f"Date: {Date}", border=0, align="L", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    raw_column_names = df.columns
    column_names = [name.replace("_", " ").title() for name in raw_column_names]

    pdf.set_font(family="Times", size=12, style="B")
    pdf.cell(w=30, h=10, txt=column_names[0], border=1)
    pdf.cell(w=60, h=10, txt=column_names[1], border=1)
    pdf.cell(w=40, h=10, txt=column_names[2], border=1)
    pdf.cell(w=30, h=10, txt=column_names[3], border=1)
    pdf.cell(w=30, h=10, txt=column_names[4], border=1, ln=1)

    for index,row in df.iterrows():
        pdf.set_font(family="Times", size=12)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=10, txt=str(row[raw_column_names[0]]), border=1)
        pdf.cell(w=60, h=10, txt=str(row[raw_column_names[1]]), border=1)
        pdf.cell(w=40, h=10, txt=str(row[raw_column_names[2]]), border=1)
        pdf.cell(w=30, h=10, txt=str(row[raw_column_names[3]]), border=1)
        pdf.cell(w=30, h=10, txt=str(row[raw_column_names[4]]), border=1, ln=1)

    Sum_Total_Price = df['total_price'].sum()

    pdf.cell(w=30, h=10, txt="", border=1)
    pdf.cell(w=60, h=10, txt="", border=1)
    pdf.cell(w=40, h=10, txt="", border=1)
    pdf.cell(w=30, h=10, txt="", border=1)
    pdf.cell(w=30, h=10, txt=str(Sum_Total_Price), border=1, ln=1)



    pdf.output(f"PDFs/{filename}.pdf")
