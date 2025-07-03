import subprocess
import streamlit as st

def convert_doc_to_pdf(docx_file, out_dir):
    subprocess.run([
        "libreoffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", out_dir,
        docx_file
    ], check=True)
    return f"{out_dir}/{docx_file.name[:-5]}.pdf"

uploaded = st.file_uploader("Upload .docx", type="docx")
if uploaded:
    with open(uploaded.name, "wb") as f:
        f.write(uploaded.getbuffer())
    pdf_path = convert_doc_to_pdf(uploaded, ".")
    st.success("Converted to PDF!")
    st.download_button("Download PDF", data=open(pdf_path, "rb"), file_name="converted.pdf")
