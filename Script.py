import os
from pathlib import Path
import pandas as pd
from docxtpl import DocxTemplate
import re
import inflect
from datetime import datetime
import win32com.client
import time
import calendar  # To get the full month name
import tkinter as tk
from tkinter import filedialog

# Initialize tkinter root for file dialog
root = tk.Tk()
root.withdraw()  # Hide the root window

# Open file dialog to select Word template file
word_template_path = filedialog.askopenfilename(title="Select Word Template File", filetypes=[("Word Files", "*.docx")])
if not word_template_path:
    print("No Word template selected. Exiting...")
    exit()

# Open file dialog to select Excel file
excel_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel Files", "*.xlsx")])
if not excel_path:
    print("No Excel file selected. Exiting...")
    exit()

# Define output directory
output_dir = Path.cwd() / "OUTPUT"
output_dir.mkdir(exist_ok=True)

# Load the Excel file
df = pd.read_excel(excel_path, sheet_name="Sheet2")
df1 = pd.read_excel(excel_path, sheet_name="Sheet1")
df = pd.merge(df, df1, on="GSTIN", how='inner')
df['Total_Amount'] = df[['CGST', 'SGST', 'IGST', 'Taxable_Value']].sum(axis=1)
df['GST_Rate'] = df[['Tax_Rate1', 'Tax_Rate_2', 'Tax_Rate_3']].sum(axis=1)
df['Invoice_Number'] = 'SMRTIPL/'+df['State_Code'].astype(str)+'/RCM/' + df['Fiscal_Period'].astype(str)+'-' + df['Fiscal_Year'].astype(str)
df['Accounting_Date'] = pd.to_datetime(df['Accounting_Date']).dt.date

p = inflect.engine()

def number_to_words_currency(num):
    if isinstance(num, float):
        rupees, paise = str(num).split(".")
        rupees_in_words = p.number_to_words(int(rupees)).capitalize()
        paise_in_words = p.number_to_words(int(paise)).capitalize()
        return f"{rupees_in_words} Rupees and {paise_in_words} Paise"
    else:
        return p.number_to_words(int(num)).capitalize() + " Rupees"

df['In_Words'] = df['Total_Amount'].apply(number_to_words_currency)

def sanitize_filename(filename):
    return re.sub(r'[\\/*?:"<>|]', '_', filename)

# Initialize dictionary for state-specific sequence numbers
state_sequence = {}

# Initialize counter for tracking progress
total_records = len(df)
counter = 0

for record in df.to_dict(orient="records"):
    fiscal_year = str(record['Fiscal_Year'])
    fiscal_period = str(record['Fiscal_Period']).zfill(2)  # Ensure fiscal period is 2 digits, e.g., '01', '02', etc.
    
    # Convert Fiscal_Period (month number) to full month name
    month_name = calendar.month_name[int(fiscal_period)]  # Get full month name (e.g., "January", "February")
    
    # Create the invoice number with the state-specific sequence counter
    state_code = record['State_Code']
    if state_code not in state_sequence:
        state_sequence[state_code] = 1
    else:
        state_sequence[state_code] += 1
    invoice_number = f"SMRTIPL/{state_code}/{record['Fiscal_Period']}-{record['Fiscal_Year']}-{state_sequence[state_code]:04d}"
    record['Invoice_Number'] = invoice_number  # Update the Invoice_Number field with the state-specific sequence number
    
    doc = DocxTemplate(word_template_path)
    doc.render(record)
    
    vendor = sanitize_filename(str(record['Vendor']))
    invoice_no = sanitize_filename(str(record['Supplier_Invoice_No']))
    
    # Get Address_3, Fiscal_Year, and Fiscal_Period (converted to full month name)
    address_3 = sanitize_filename(str(record['Address_3']))
    
    # Create the directory structure based on Address_3, Fiscal_Year, and full month name for Fiscal_Period
    address_3_dir = output_dir / address_3  # Address_3 directory
    fiscal_year_dir = address_3_dir / fiscal_year  # Fiscal Year directory
    fiscal_period_dir = fiscal_year_dir / month_name  # Fiscal Period as full month name (e.g., "January")
    
    # Create all directories if they don't exist
    fiscal_period_dir.mkdir(parents=True, exist_ok=True)
    
    # Set the output path for the PDF
    pdf_output_path = fiscal_period_dir / f"{fiscal_year}_{month_name}_{vendor}_{invoice_no}.pdf"
    
    # Save the document as .docx
    doc.save(pdf_output_path.with_suffix(".docx"))

    # Convert the .docx to .pdf using Word
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(str(pdf_output_path.with_suffix(".docx")))
    
    # Save as PDF
    doc.SaveAs(str(pdf_output_path), FileFormat=17)  # 17 is the constant for PDF format
    
    # Wait for the save to finish
    time.sleep(1)  # Add a small delay to ensure the file is saved properly
    
    # Close the Word document
    doc.Close()
    
    # Quit Word application
    word.Quit()

    # Optionally, delete the .docx file after saving the PDF
    try:
        os.remove(pdf_output_path.with_suffix(".docx"))
    except PermissionError:
        print(f"PermissionError: The file {pdf_output_path.with_suffix('.docx')} is still in use and could not be deleted.")
    
    # Increment the counter and print the progress
    counter += 1
    os.system('cls')
    print(f"Converted and saved {counter}/{total_records} PDFs.")

print("Conversion to PDF completed for all documents.")
