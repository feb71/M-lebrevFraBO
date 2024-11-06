import fitz  # PyMuPDF
import re
import os
import zipfile
from datetime import datetime
from pathlib import Path
import streamlit as st

st.set_page_config(layout="wide")  # Bruk hele bredden av skjermen

# Funksjon for å kombinere hoved-PDF og vedlegg
def combine_pdf_and_attachments(pdf_file, folder_files):
    combined_document = fitz.open()
    original_document = fitz.open(stream=pdf_file.read(), filetype="pdf")
    folder_dict = {Path(file.name).name: file for file in folder_files}

    for page_num in range(len(original_document)):
        page = original_document.load_page(page_num)
        combined_document.insert_pdf(original_document, from_page=page_num, to_page=page_num)
        text = page.get_text("text")
        links_text = text.split("Vedlagte dokumenter:")[1].strip().split("\n") if "Vedlagte dokumenter:" in text else []

        for link_text in links_text:
            link_text = link_text.strip().replace("\\", "/").split("/")[-1]
            if not re.match(r'.+\.pdf$', link_text):
                continue
            if link_text in folder_dict:
                attachment_file = folder_dict[link_text]
                attachment_file.seek(0)
                attachment_document = fitz.open(stream=attachment_file.read(), filetype="pdf")
                combined_document.insert_pdf(attachment_document)
                attachment_document.close()

    output_path = 'kombinert_dokument.pdf'
    combined_document.save(output_path)
    combined_document.close()

    return output_path

# Funksjon for å trekke ut verdier fra teksten
def trekk_ut_verdier(tekst):
    # Regex for å finne postnummeret basert på "Postnummer Beskrivelse" og enhetstype
    postnummer_pattern = r'postnummer\s+beskrivelse\s+([\d.]+)\s+(?=rs|stk|kg|m|m2|m3)\b'
    postnummer_match = re.search(postnummer_pattern, tekst, re.IGNORECASE)

    postnummer = postnummer_match.group(1).strip() if postnummer_match else "ukjent"
    st.write(f"Funnet postnummer: {postnummer}")  # Debug for å bekrefte funnet postnummer

    # Finn mengde og dato som før
    mengde_pattern = r'(?<=Utført pr. d.d.:\n)([\d,]+)'
    dato_pattern = r'(\d{2}\.\d{2}\.\d{4})'

    mengde_match = re.search(mengde_pattern, tekst)
    dato_match = datetime.now().strftime("%Y%m%d")

    mengde = mengde_match.group(1) if mengde_match else "ukjent"

    if dato_match := re.search(dato_pattern, tekst):
        dato_match = datetime.strptime(dato_match.group(1), "%d.%m.%Y").strftime("%Y%m%d")

    return postnummer, mengde, dato_match

# Funksjon for å opprette nye PDF-er
def opprett_ny_pdf(original_pdf, startside, sluttside, output_path):
    original_pdf.seek(0)
    dokument = fitz.open(stream=original_pdf.read(), filetype="pdf")
    ny_pdf = fitz.open()
    ny_pdf.insert_pdf(dokument, from_page=startside, to_page=sluttside)
    ny_pdf.save(output_path)
    ny_pdf.close()
    dokument.close()

# Funksjon for å lese tekst fra PDF for splitting
def les_tekst_fra_pdf(pdf_file):
    dokument = fitz.open(stream=pdf_file.read(), filetype="pdf")
    tekst_per_side = []
    for side_num in range(len(dokument)):
        side = dokument[side_num]
        tekst_per_side.append(side.get_text("text"))
    dokument.close()
    return tekst_per_side

# Funksjon for å zippe en mappe
def zip_directory(directory_path, output_zip_path):
    with zipfile.ZipFile(output_zip_path, 'w') as zipf:
        for root, dirs, files in os.walk(directory_path):
            for file in files:
                file_path = os.path.join(root, file)
                zipf.write(file_path, os.path.relpath(file_path, directory_path))

# Funksjon for å behandle og splitte PDF
def behandle_og_splitte_pdf(uploaded_pdf, output_folder):
    tekst_per_side = les_tekst_fra_pdf(uploaded_pdf)
    startside = 0
    opprettede_filer = []

    for i, tekst in enumerate(tekst_per_side):
        if "Målebrev" in tekst and i > startside:
            postnummer, mengde, dato = trekk_ut_verdier(tekst_per_side[startside])
            filnavn = f"{postnummer}_{dato}.pdf"
            output_sti = os.path.join(output_folder, filnavn)
            uploaded_pdf.seek(0)
            st.write(f"Oppretter PDF: {output_sti} fra side {startside} til {i - 1}")
            opprett_ny_pdf(uploaded_pdf, startside, i - 1, output_sti)
            opprettede_filer.append(output_sti)
            startside = i

    # Håndter siste segment
    postnummer, mengde, dato = trekk_ut_verdier(tekst_per_side[startside])
    filnavn = f"{postnummer}_{dato}.pdf"
    output_sti = os.path.join(output_folder, filnavn)
    uploaded_pdf.seek(0)
    st.write(f"Oppretter siste PDF: {output_sti} fra side {startside} til {len(tekst_per_side) - 1}")
    opprett_ny_pdf(uploaded_pdf, startside, len(tekst_per_side) - 1, output_sti)
    opprettede_filer.append(output_sti)

    return opprettede_filer

# Streamlit-grensesnittet
st.title("Kombiner målebrev med Vedlegg / Lag en fil pr post")
col1, col2 = st.columns(2)

# Kombiner målebrev med vedlegg
with col1:
    st.subheader("Kombiner målebrev med Vedlegg")
    pdf_file = st.file_uploader("Last opp PDF-filen med Målebrev", type="pdf", key="combine_pdf")
    folder_files = st.file_uploader("Last opp vedleggs-PDF-filer", type="pdf", accept_multiple_files=True, key="attachments")

    if pdf_file and folder_files:
        output_path = combine_pdf_and_attachments(pdf_file, folder_files)
        with open(output_path, "rb") as f:
            st.download_button("Last ned kombinert PDF", f, file_name="kombinert_dokument.pdf")

# Splitt PDF pr post
with col2:
    st.subheader("Splitt PDF-fil pr post")
    uploaded_pdf = st.file_uploader("Last opp PDF-fil for splitting", type="pdf", key="split_pdf")
    output_folder = os.path.join(os.path.expanduser("~"), "Downloads", "Splittet_malebrev")

    if uploaded_pdf and st.button("Start Splitting av PDF", key="split_button"):
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        
        st.write(f"Starter splitting og lagrer i mappe: {output_folder}")
        opprettede_filer = behandle_og_splitte_pdf(uploaded_pdf, output_folder)

        zip_path = os.path.join(os.path.expanduser("~"), "Downloads", "Splittet_malebrev.zip")
        zip_directory(output_folder, zip_path)

        with open(zip_path, "rb") as z:
            st.download_button("Last ned alle PDF-filer som ZIP", z, file_name="Splittet_malebrev.zip", mime="application/zip")
