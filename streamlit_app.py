import fitz  # PyMuPDF
import re
import os
import zipfile
from datetime import datetime
from pathlib import Path
import streamlit as st

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
    dato_match = datetime.now().strftime("%Y%m%d")
    
    if dato_match := re.search(dato_pattern, tekst):
        dato_match = datetime.strptime(dato_match.group(1), "%d.%m.%Y").strftime("%Y%m%d")
    
    postnummer = postnummer_match.group(1) if postnummer_match else "ukjent"
    mengde = mengde_match.group(1) if mengde_match else "ukjent"

    return postnummer, mengde, dato_match

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
st.set_page_config(layout="wide")  # Bruk hele bredden av skjermen
st.title("Splitt PDF pr post")

# Kolonne 1: Velg handlinger og nedlastingsknapper
med_splitting = st.checkbox("Splitt PDF pr post")

# Kolonne for opplasting og splitting
if med_splitting:
    st.subheader("Splitt PDF-fil pr post")
    uploaded_pdf = st.file_uploader("Last opp PDF-fil for splitting", type=["pdf"], key="split_pdf")
    
    if uploaded_pdf and st.button("Start Splitting av PDF"):
        ny_mappe = os.path.join(os.path.expanduser("~"), "Downloads", "Splittet_malebrev")
        if not os.path.exists(ny_mappe):
            os.makedirs(ny_mappe)

        tekst_per_side = les_tekst_fra_pdf(uploaded_pdf)
        opprettede_filer = []
        startside = 0
        for i, tekst in enumerate(tekst_per_side):
            # Splitter målebrev ved "Målebrev" i teksten, uavhengig av om vedlegg er tilstede
            if "Målebrev" in tekst and i > startside:
                postnummer, mengde, dato = trekk_ut_verdier(tekst_per_side[startside])
                filnavn = f"{postnummer}_{dato}.pdf"
                output_sti = os.path.join(ny_mappe, filnavn)
                uploaded_pdf.seek(0)
                opprett_ny_pdf(uploaded_pdf, startside, i - 1, output_sti)
                opprettede_filer.append(output_sti)
                startside = i

        # Sørg for at siste målebrev alltid splittes
        postnummer, mengde, dato = trekk_ut_verdier(tekst_per_side[startside])
        filnavn = f"{postnummer}_{dato}.pdf"
        output_sti = os.path.join(ny_mappe, filnavn)
        uploaded_pdf.seek(0)
        opprett_ny_pdf(uploaded_pdf, startside, len(tekst_per_side) - 1, output_sti)
        opprettede_filer.append(output_sti)

        # Opprett zip-fil av splittede PDF-er
        zip_filnavn = os.path.join(os.path.expanduser("~"), "Downloads", "Splittet_malebrev.zip")
        zip_directory(ny_mappe, zip_filnavn)

        # Tilby zip-filen for nedlasting
        with open(zip_filnavn, "rb") as z:
            st.download_button(
                label="Last ned alle PDF-filer som ZIP",
                data=z,
                file_name="Splittet_malebrev.zip",
                mime="application/zip"
            )

