import platform
import os
import re
import calendar
import tempfile
import zipfile
import subprocess
from io import BytesIO
from pathlib import Path
from datetime import datetime

import pandas as pd
import inflect
import streamlit as st
from docxtpl import DocxTemplate

# Try to import PDF converters
try:
    from docx2pdf import convert as docx2pdf_convert
except ImportError:
    docx2pdf_convert = None

import pypandoc

# --- Pandoc Setup for Linux/Streamlit Cloud ---
def setup_pandoc():
    """Set up pandoc manually if not found (for Linux/Streamlit Cloud)."""
    try:
        if not pypandoc.get_pandoc_path():
            pandoc_dir = Path("pandoc")
            pandoc_dir.mkdir(exist_ok=True)
            deb_path = pandoc_dir / "pandoc.deb"
            url = "https://github.com/jgm/pandoc/releases/download/3.1.11.1/pandoc-3.1.11.1-1-amd64.deb"

            subprocess.run(["wget", "-O", str(deb_path), url], check=True)
            subprocess.run(["dpkg", "-x", str(deb_path), str(pandoc_dir)], check=True)

            bin_path = pandoc_dir / "usr/bin"
            os.environ["PATH"] = f"{bin_path}:{os.environ['PATH']}"
            pypandoc.get_pandoc_path()  # Confirm it's available
        return True
    except Exception as e:
        st.error(f"Failed to set up pandoc: {e}")
        return False

def convert_docx_to_pdf_pandoc(docx_path, pdf_output_path):
    try:
        output = pypandoc.convert_file(str(docx_path), 'pdf', outputfile=str(pdf_output_path))
        return True
    except Exception as e:
        st.error(f"PDF conversion (pandoc) failed: {e}")
        return False

# --- STREAMLIT UI ---

st.set_page_config(page_title="Invoice Generator", layout="wide")
st.title("📄 Automated Invoice Generator with PDF Export")

st.markdown("Upload your Word Template and Excel File to begin:")

word_template_file = st.file_uploader("Upload Word Template (.docx)", type=["docx"])
excel_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])

if word_template_file and excel_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_word:
        tmp_word.write(word_template_file.read())
        word_template_path = tmp_word.name

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_excel:
        tmp_excel.write(excel_file.read())
        excel_path = tmp_excel.name

    # --- PROCESSING ---

    output_dir = Path.cwd() / "OUTPUT"
    output_dir.mkdir(exist_ok=True)

    df = pd.read_excel(excel_path, sheet_name="Sheet2")
    df1 = pd.read_excel(excel_path, sheet_name="Sheet1")
    df = pd.merge(df, df1, on="GSTIN", how='inner')
    df['Total_Amount'] = df[['CGST', 'SGST', 'IGST', 'Taxable_Value']].sum(axis=1)
    df['GST_Rate'] = df[['Tax_Rate1', 'Tax_Rate_2', 'Tax_Rate_3']].sum(axis=1)
    df['Invoice_Number'] = 'SMRTIPL/' + df['State_Code'].astype(str) + '/RCM/' + df['Fiscal_Period'].astype(str) + '-' + df['Fiscal_Year'].astype(str)
    df['Accounting_Date'] = pd.to_datetime(df['Accounting_Date']).dt.date

    p = inflect.engine()

    def number_to_words_currency(num):
        if isinstance(num, float):
            rupees, paise = str(num).split(".")
            return f"{p.number_to_words(int(rupees)).capitalize()} Rupees and {p.number_to_words(int(paise)).capitalize()} Paise"
        else:
            return p.number_to_words(int(num)).capitalize() + " Rupees"

    df['In_Words'] = df['Total_Amount'].apply(number_to_words_currency)

    def sanitize_filename(filename):
        return re.sub(r'[\\/*?:"<>|]', '_', filename)

    state_sequence = {}
    total_records = len(df)
    progress = st.progress(0)
    status_text = st.empty()

    # Setup pandoc if needed
    if platform.system() not in ['Windows', 'Darwin']:
        setup_pandoc()

    for counter, record in enumerate(df.to_dict(orient="records"), start=1):
        fiscal_year = str(record['Fiscal_Year'])
        fiscal_period = str(record['Fiscal_Period']).zfill(2)
        month_name = calendar.month_name[int(fiscal_period)]

        state_code = record['State_Code']
        state_sequence[state_code] = state_sequence.get(state_code, 0) + 1
        invoice_number = f"SMRTIPL/{state_code}/{fiscal_period}-{fiscal_year}-{state_sequence[state_code]:04d}"
        record['Invoice_Number'] = invoice_number

        doc = DocxTemplate(word_template_path)
        doc.render(record)

        vendor = sanitize_filename(str(record['Vendor']))
        invoice_no = sanitize_filename(str(record['Supplier_Invoice_No']))
        address_3 = sanitize_filename(str(record['Address_3']))

        address_3_dir = output_dir / address_3 / fiscal_year / month_name
        address_3_dir.mkdir(parents=True, exist_ok=True)

        pdf_output_path = address_3_dir / f"{fiscal_year}_{month_name}_{vendor}_{invoice_no}.pdf"
        docx_path = pdf_output_path.with_suffix(".docx")
        doc.save(docx_path)

        # --- Convert DOCX to PDF ---
        if platform.system() in ['Windows', 'Darwin'] and docx2pdf_convert:
            try:
                docx2pdf_convert(str(docx_path), str(pdf_output_path))
                os.remove(docx_path)
            except Exception as e:
                st.error(f"docx2pdf failed: {e}")
        else:
            if convert_docx_to_pdf_pandoc(docx_path, pdf_output_path):
                os.remove(docx_path)
            else:
                st.warning("PDF conversion failed. DOCX saved instead.")

        progress.progress(counter / total_records)
        status_text.text(f"Generated {counter}/{total_records} invoices.")

    st.success("✅ All invoices generated and saved as PDFs in the OUTPUT folder.")

    # --- ZIP and Download ---
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for folder_path, _, files in os.walk(output_dir):
            for file in files:
                if file.endswith(".pdf"):
                    file_path = os.path.join(folder_path, file)
                    arcname = os.path.relpath(file_path, output_dir)
                    zipf.write(file_path, arcname=arcname)
    zip_buffer.seek(0)

    st.download_button(
        label="📥 Download All PDFs as ZIP",
        data=zip_buffer,
        file_name="invoices.zip",
        mime="application/zip"
    )

else:
    st.info("Please upload both Word and Excel files to proceed.")
