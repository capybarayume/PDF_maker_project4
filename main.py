import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

class PDF(FPDF):
    def footer(self):
        self.set_y(-15)
        # Setting font: helvetica italic 8
        self.set_font(family="helvetica",style= "I", size=8)
        # Printing page number:
        self.cell(w=0, h=10, text= "Page {0}/{1}".format(self.page_no(),"{nb}"), align="C")


#using invoice as a example

filepaths = glob.glob("input/*.xlsx", recursive=True)

for filepath in filepaths:
    pdf=PDF(orientation="P", unit="mm", format="A4")    

    pdf.add_page()

    filename=Path(filepath).stem
    invoice_nr=filename.split("-")


    pdf.set_font(family="Times",size=16,style="B")
    pdf.cell(w=50,h=8,text="Invoice.nr.{}".format(invoice_nr[0]),ln=1)

    pdf.set_font(family="Times",size=16,style="B")
    pdf.cell(w=50,h=8,text="Date: {}".format(invoice_nr[1]),ln=1)

    df=pd.read_excel(filepath,sheet_name="Sheet 1")
    headers = df.columns.to_list()
    data=df.to_numpy()
    

    #header
    headers=[header.replace("_"," ") for header in headers]
    pdf.set_font(family="Times",size=10)
    pdf.cell(w=30,h=8,text=headers[0],border=1)
    pdf.cell(w=70,h=8,text=headers[1],border=1)
    pdf.cell(w=30,h=8,text=headers[2],border=1)
    pdf.cell(w=30,h=8,text=headers[3],border=1)
    pdf.cell(w=30,h=8,text=headers[4],border=1,ln=1)

    #form data
    for row in data:
        pdf.set_font(family="Times",size=10)
        pdf.set_text_color(80,80,80)
        pdf.cell(w=30,h=8,text=str(row[0]),border=1)
        pdf.cell(w=70,h=8,text=str(row[1]),border=1)
        pdf.cell(w=30,h=8,text=str(row[2]),border=1)
        pdf.cell(w=30,h=8,text=str(row[3]),border=1)
        pdf.cell(w=30,h=8,text=str(row[4]),border=1,ln=1)

    #total price box
    price_sum=data[:,4].sum()
    pdf.set_font(family="Times",size=10)
    pdf.set_text_color(80,80,80)
    pdf.cell(w=30,h=8,text="",border=1)
    pdf.cell(w=70,h=8,text="",border=1)
    pdf.cell(w=30,h=8,text="",border=1)
    pdf.cell(w=30,h=8,text="",border=1)
    pdf.cell(w=30,h=8,text=str(price_sum),border=1,ln=1)

    #sentences total sum
    pdf.set_font(family="Times",size=10,style="B")
    pdf.cell(w=30,h=8,text="The total price is {}".format(price_sum),ln=1)

    # add company logo 
    pdf.set_font(family="Times",size=14,style="B")
    pdf.multi_cell(w=25,h=8,text="website",)


    pdf.output("input/PDFs/{0}.pdf".format(filename))
    
