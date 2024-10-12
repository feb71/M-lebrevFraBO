import fitz  # PyMuPDF
import re
import os
import zipfile
from datetime import datetime
from pathlib import Path
import streamlit as st

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
                st.info(f"Legger til vedlegget: {link_text} for side {page_num + 1}")
                attachment_file.seek(0)
                attachment_document = fitz.open(stream=attachment_file.read(), filetype="pdf")
                combined_document.insert_pdf(attachment_document)
                attachment_document.close()
            else:
                st.warning(f"Fant ikke vedlegget: {link_text} i opplastede filer")

    output_path = 'kombinert_dokument.pdf'
    combined_document.save(output_path)
    combined_document.close()

    return output_path

# Funksjon for å lese tekst fra PDF for splitting
def les_tekst_fra_pdf(pdf_file):
    dokument = fitz.open(stream=pdf_file.read(), filetype="pdf")
    tekst_per_side = []
    for side_num in range(len(dokument)):
        side = dokument[side_num]
        tekst_per_side.append(side.get_text("text"))
    dokument.close()
    return tekst_per_side

# Funksjon for å trekke ut verdier fra teksten
def trekk_ut_verdier(tekst):
    beskrivelse_pattern = r'Beskrivelse\s*(\d{1,2}\.\d{1,2}\.\d{2,3})'
    mengde_pattern = r'(?<=Utført pr. d.d.:\n)([\d,]+)'
    dato_pattern = r'(\d{2}\.\d{2}\.\d{4})'

    postnummer_match = re.search(beskrivelse_pattern, tekst)
    mengde_match = re.search(mengde_pattern, tekst)
    dato_match = re.search(dato_pattern, tekst)

    postnummer = postnummer_match.group(1) if postnummer_match else "ukjent"
    mengde = mengde_match.group(1) if mengde_match else "ukjent"
    dato = datetime.strptime(dato_match.group(1), "%d.%m.%Y").strftime("%Y%m%d") if dato_match else datetime.now().strftime("%Y%m%d")

    return postnummer, mengde, dato

def opprett_ny_pdf(original_pdf, startside, sluttside, output_path):
    original_pdf.seek(0)
    dokument = fitz.open(stream=original_pdf.read(), filetype="pdf")
    ny_pdf = fitz.open()
    ny_pdf.insert_pdf(dokument, from_page=startside, to_page=sluttside)
    ny_pdf.save(output_path)
    ny_pdf.close()
    dokument.close()

def zip_directory(directory_path, output_zip_path):
    with zipfile.ZipFile(output_zip_path, 'w') as zipf:
        for root, dirs, files in os.walk(directory_path):
            for file in files:
                file_path = os.path.join(root, file)
                zipf.write(file_path, os.path.relpath(file_path, directory_path))

# Streamlit app
st.title("PDF Generator og Splitting")

# Opprett to kolonner (side om side)
col1, col2 = st.columns(2)

# Bruker velger om de vil generere PDF, splitte PDF, eller begge deler
med_generering = st.checkbox("Kombiner PDF med vedlegg")
med_splitting = st.checkbox("Splitt PDF")

# Finn brukermappen
brukermappe = os.path.expanduser("~")
downloads_mappe = os.path.join(brukermappe, "Downloads")

# Generer PDF (kolonne 1)
if med_generering:
    with col1:
        st.subheader("Kombiner PDF med Vedlegg")
        pdf_file = st.file_uploader("Last opp PDF-filen med Målebrev", type="pdf")
        folder_files = st.file_uploader("Last opp vedleggs-PDF-filer", type="pdf", accept_multiple_files=True)
        
        if pdf_file is not None and folder_files:
            st.write("Kombinerer filene, vennligst vent...")
            output_path = combine_pdf_and_attachments(pdf_file, folder_files)
            st.success("Kombinering fullført!")
            with open(output_path, "rb") as f:
                st.download_button("Last ned kombinert PDF", f, file_name="kombinert_dokument.pdf")

# Splitt PDF (kolonne 2)
if med_splitting:
    with col2:
        st.subheader("Splitt en PDF-fil")
        uploaded_pdf = st.file_uploader("Last opp PDF-fil for splitting", type=["pdf"])
        
        if uploaded_pdf and st.button("Start Splitting av PDF"):
            ny_mappe = os.path.join(downloads_mappe, "Splittet_malebrev")
            if not os.path.exists(ny_mappe):
                os.makedirs(ny_mappe)

            tekst_per_side = les_tekst_fra_pdf(uploaded_pdf)
            opprettede_filer = []
            startside = 0
            for i, tekst in enumerate(tekst_per_side):
                if "Målebrev" i tekst and i > startside:
                    postnummer, mengde, dato = trekk_ut_verdier(tekst_per_side[startside])
                    filnavn = f"{postnummer}_{dato}.pdf"
                    output_sti = os.path.join(ny_mappe, filnavn)
                    uploaded_pdf.seek(0)
                    opprett_ny_pdf(uploaded_pdf, startside, i - 1, output_sti)
                    opprettede_filer.append(output_sti)
                    startside = i

            postnummer, mengde, dato = trekk_ut_verdier(tekst_per_side[startside])
            filnavn = f"{postnummer}_{dato}.pdf"
            output_sti = os.path.join(ny_mappe, filnavn)
            uploaded_pdf.seek(0)
            opprett_ny_pdf(uploaded_pdf, startside, len(tekst_per_side) - 1, output_sti)
            opprettede_filer.append(output_sti)

            zip_filnavn = os.path.join(downloads_mappe, "Splittet_malebrev.zip")
            zip_directory(ny_mappe, zip_filnavn)

            with open(zip_filnavn, "rb") as z:
                st.download_button(
                    label="Last ned alle PDF-filer som ZIP",
                    data=z,
                    file_name="Splittet_malebrev.zip",
                    mime="application/zip"
                )
