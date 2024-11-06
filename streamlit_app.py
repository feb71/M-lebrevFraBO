import fitz  # PyMuPDF
import re
import os
import zipfile
from datetime import datetime
from pathlib import Path
import streamlit as st

st.set_page_config(layout="wide")  # Bruk hele bredden av skjermen

# Funksjon for å trekke ut verdier fra teksten
def trekk_ut_verdier(tekst):
    postnummer_pattern = r'postnummer\s+beskrivelse\s+([\d.]+)\s+(?=rs|stk|kg|m|m2|m3)\b'
    postnummer_match = re.search(postnummer_pattern, tekst, re.IGNORECASE)
    postnummer = postnummer_match.group(1).strip() if postnummer_match else "ukjent"
    
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
            opprett_ny_pdf(uploaded_pdf, startside, i - 1, output_sti)
            opprettede_filer.append(output_sti)
            startside = i

    postnummer, mengde, dato = trekk_ut_verdier(tekst_per_side[startside])
    filnavn = f"{postnummer}_{dato}.pdf"
    output_sti = os.path.join(output_folder, filnavn)
    uploaded_pdf.seek(0)
    opprett_ny_pdf(uploaded_pdf, startside, len(tekst_per_side) - 1, output_sti)
    opprettede_filer.append(output_sti)

    return opprettede_filer

# Funksjon for å kombinere hovedmålebrev med vedleggene
def combine_pdf_and_attachments(pdf_file, folder_files, output_path):
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

    combined_document.save(output_path)
    combined_document.close()

# Funksjon for å zippe kun de opprettede filene
def zip_opprettede_filer(filer, zip_filnavn):
    with zipfile.ZipFile(zip_filnavn, 'w') as zipf:
        for fil in filer:
            zipf.write(fil, os.path.basename(fil))

# Streamlit-grensesnittet
st.title("Kombiner målebrev med vedlegg / Lag en fil pr post")
col1, col2, col3 = st.columns([1, 2, 2])

# Velg handlinger
with col1:
    st.write("## Velg handlinger")
    med_generering = st.checkbox("Kombiner målebrev med vedlegg")
    med_splitting = st.checkbox("Splitt kombinert PDF pr post")

output_folder = os.path.join(os.path.expanduser("~"), "Downloads", "Splittet_malebrev")

# Kombiner PDF og vedlegg
if med_generering:
    with col2:
        st.subheader("Kombiner målebrev med Vedlegg")
        pdf_file = st.file_uploader("Last opp PDF-filen med Målebrev", type="pdf", key="combine_pdf")
        with st.expander("Last opp vedleggs-PDF-filer"):
            folder_files = st.file_uploader("Velg vedleggsfiler (PDF)", type="pdf", accept_multiple_files=True, key="attachments")
        
        if pdf_file and folder_files:
            output_path = os.path.join(output_folder, "kombinert_dokument.pdf")
            combine_pdf_and_attachments(pdf_file, folder_files, output_path)
            st.success("Kombinering fullført!")
            with open(output_path, "rb") as f:
                st.download_button("Last ned kombinert PDF", f, file_name="kombinert_dokument.pdf")

# Splitt PDF pr post
if med_splitting:
    with col3:
        st.subheader("Splitt PDF-fil pr post")
        uploaded_pdf = st.file_uploader("Last opp PDF-fil for splitting", type="pdf", key="split_pdf")

        if uploaded_pdf and st.button("Start Splitting av PDF"):
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)

            opprettede_filer = behandle_og_splitte_pdf(uploaded_pdf, output_folder)

            st.success("Splitting fullført!")
            st.write(f"Antall opprettede filer: {len(opprettede_filer)}")

            # Lag en ZIP-fil med kun de opprettede filene
            zip_filnavn = os.path.join(os.path.expanduser("~"), "Downloads", "Splittet_malebrev.zip")
            zip_opprettede_filer(opprettede_filer, zip_filnavn)

            # Tilby nedlasting av ZIP-filen
            if os.path.exists(zip_filnavn):
                with open(zip_filnavn, "rb") as z:
                    st.download_button("Last ned alle PDF-filer som ZIP", z, file_name="Splittet_malebrev.zip", mime="application/zip")
            else:
                st.error("ZIP-filen ble ikke opprettet.")
