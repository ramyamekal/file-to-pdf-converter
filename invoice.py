import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Excel file to PDF
filepaths=glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df=pd.read_excel(filepath,sheet_name="Sheet 1") # this should be same name based on Excel
    pdf=FPDF(orientation="p",unit="mm",format="A4")
    pdf.add_page()
    filename= Path(filepath).stem  #it will give filename without extension
    invoice_nr,date=filename.split("-")
    pdf.set_font(family="Times",size=16,style="B")
    pdf.cell(w=50,h=8, txt=f"Invoice Nr. {invoice_nr }" , ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date. {date}",ln=1)
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    #Add Header
    columns = list(df.columns)
    columns=[item.replace("-","").title() for item in columns]
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    #Add TABLE Data
    for index,row in df.iterrows():
        total_sum=df["total_price"].sum()
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80,80,80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]),border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]),border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]),border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]),border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]),border=1,ln=1)
        pdf.cell(w=30,h=8,txt=str(total_sum),border=1,ln=1)

#Add Total Value
        pdf.set_font(family="Times", size=10)
        pdf.cell(w=30, h=8, txt=f"The total sum id{total_sum}", ln=1)

        # Add Company and logo
        pdf.set_font(family="Times", size=10, style="B")
        pdf.cell(w=25, h=8, txt=f"pythonhow")
        pdf.image("pythonhow.png",w=10)

        pdf.output(f"PDFs/{filename}.pdf")


# Text File to PDF
filepaths=glob.glob("textfiles/*.txt")
for filepath in filepaths:
    pdf=FPDF(orientation="p",unit="mm",format="A4")
    pdf.add_page()
    filename = Path(filepath).stem  # it will give filename without extension
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"{filename.title()}",ln=1)
    with open(f"textfiles/{filename}.txt", 'r') as file:
        content=file.read()
        pdf.set_font(family="Times", size=30)
        pdf.multi_cell(w=0, h=10, txt=content)

    pdf.output(f"TextToPDF/{filename}.pdf")





