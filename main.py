from tkinter.tix import COLUMN
import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob('invoices/*.xlsx')

for filepath in filepaths:
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    
    filename = Path(filepath).stem
    invoice_nr, invoice_date = filename.split('-')
    
    pdf.set_font(family='Times', style='B', size=24)
    pdf.cell(w=50, h=8, txt=f'Invoice nr.{invoice_nr}, ln=1')
    
    pdf.set_font(family='Times', style='B', size=24)
    pdf.cell(w=50, h=8, txt=f'Date nr.{invoice_date}, ln=1')
    
    df = pd.read_excel(filepath, sheet_name='Sheet 1')
    
    columns = list(df.columns)
    
    for index, row in df.iterrows():
        pdf.set_font(family='Times', size=12)
        pdf.set_text_color(r=80, g=80, b=80)
        for column in columns:
            pdf.cell(w=30, h=8, txt=str(row[column]), border=1)
        pdf.ln()
    
    total_sum = df['total_price'].sum()
    pdf.set_font(family='Times', size=12)
    pdf.set_text_color(r=80, g=80, b=80)
    pdf.cell(w=30, h=8, txt='', border=1)
    pdf.cell(w=70, h=8, txt='', border=1)
    pdf.cell(w=30, h=8, txt='', border=1)
    pdf.cell(w=30, h=8, txt='', border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1) 
    
    pdf.output(f'pdfs/{filename}.pdf')
        

    for index, row in df.iterrows():
        pdf.set_font(family='Times', size=12)
        pdf.set_text_color(r=80, g=80, b=80)
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=1)
        pdf.cell(w=70, h=8, txt=str(row['product_name']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['amount_purchased']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['total_price']), border=1, ln=1)
        
        total_sum = df['total_price'].sum()
        pdf.set_font(family='Times', size=12)
        pdf.set_text_color(r=80, g=80, b=80)
        pdf.cell(w=30, h=8, txt='', border=1)
        pdf.cell(w=70, h=8, txt='', border=1)
        pdf.cell(w=30, h=8, txt='', border=1)
        pdf.cell(w=30, h=8, txt='', border=1)
        pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1) 
        
        
        # Add total sum
        pdf.set_font(family='Times', size=12, style='B')
        pdf.cell(w=30, h=8, txt=f'The total price is {total_sum}', ln=1)
        
        
        
        # Add company name and logo
        pdf.set_font(family='Times', size=14, style='B')
        pdf.cell(w=30, h=8, txt=f'Company name',) # Add company name here
        pdf.image() # Add logo here
          
        

    
    
    pdf.output(f'pdfs/{filename}.pdf')
