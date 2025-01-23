import fitz  # PyMuPDF
import re
from datetime import datetime
from pathlib import Path
import os
import zipfile
import streamlit as st
from io import BytesIO

# Funksjon for å trekke ut verdier fra teksten
def trekk_ut_verdier(tekst):
    postnummer_pattern = r'postnummer\s+beskrivelse\s+([\d.]+)\s+(?=rs|stk|kg|m|m2|m3)\b'
    postnummer_match = re.search(postnummer_pattern, tekst, re.IGNORECASE)
    postnummer = postnummer_match.group(1).strip() if postnummer_match else "ukjent"

    mengde_pattern = r'(?<=Utført pr. d.d.:\n)([\d,]+)'
    dato_pattern = r'(\d{2}\.\d{2}\.\d{4})'

    # Forsøk å finne mengde fra teksten
    mengde_match = re.search(mengde_pattern, tekst)
    mengde = mengde_match.group(1) if mengde_match else "ukjent"

    # Dato: først default = dagens dato i format YYYYMMDD
    dato_match_str = datetime.now().strftime("%Y%m%d")
    # Dersom vi faktisk finner en dato i teksten:
    if (dato_re_match := re.search(dato_pattern, tekst)):
        dato_match_str = datetime.strptime(dato_re_match.group(1), "%d.%m.%Y").strftime("%Y%m%d")

    return postnummer, mengde, dato_match_str

# Funksjon for å kombinere hovedmålebrev med vedleggene
def combine_pdf_and_attachments(pdf_file, folder_files):
    combined_document = fitz.open()
    original_document = fitz.open(stream=pdf_file.read(), filetype="pdf")
    folder_dict = {Path(file.name).name: file for file in folder_files}

    for page_num in range(len(original_document)):
        page = original_document.load_page(page_num)
        combined_document.insert_pdf(original_document, from_page=page_num, to_page=page_num)
        text = page.get_text("text")

        # Finn 'Vedlagte dokumenter:' og skill ut linjene
        if "Vedlagte dokumenter:" in text:
            links_text = text.split("Vedlagte dokumenter:")[1].strip().split("\n")
        else:
            links_text = []

        # For hver PDF-link i teksten: sjekk om filen finnes i opplastede vedlegg, og legg den til
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

    output_pdf = BytesIO()
    combined_document.save(output_pdf)
    combined_document.close()
    output_pdf.seek(0)
    return output_pdf

# Funksjon for å lese all tekst fra PDF og få en liste av tekster pr. side
def les_tekst_fra_pdf(pdf_file):
    pdf_file.seek(0)
    dokument = fitz.open(stream=pdf_file.read(), filetype="pdf")
    tekst_per_side = []
    for side_num in range(len(dokument)):
        side = dokument[side_num]
        tekst_per_side.append(side.get_text("text"))
    dokument.close()
    return tekst_per_side

# Funksjon for å opprette en ny PDF fra utvalgte sider av originalen
def opprett_ny_pdf(original_pdf, startside, sluttside, output_pdf):
    original_pdf.seek(0)
    dokument = fitz.open(stream=original_pdf.read(), filetype="pdf")
    ny_pdf = fitz.open()
    ny_pdf.insert_pdf(dokument, from_page=startside, to_page=sluttside)
    ny_pdf.save(output_pdf)
    ny_pdf.close()
    dokument.close()

# Funksjon for å splitte PDF-fil pr. post og lagre i en ZIP
def split_pdf_to_zip(combined_pdf):
    # Les tekst for hver side
    tekst_per_side = les_tekst_fra_pdf(combined_pdf)
    startside = 0
    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for i in range(len(tekst_per_side)):
            # Hvis vi finner "Målebrev" og ikke er på første side, er det starten på neste målebrev
            if "Målebrev" in tekst_per_side[i] and i > startside:
                # Samle teksten fra *hele målebrevet* (startside til i-1)
                tekst_for_ett_malebrev = "\n".join(tekst_per_side[startside:i])
                postnummer, mengde, dato = trekk_ut_verdier(tekst_for_ett_malebrev)

                filnavn = f"{postnummer}_{dato}.pdf"
                output_pdf = BytesIO()
                # Lag PDF fra sidene startside -> i-1
                opprett_ny_pdf(combined_pdf, startside, i - 1, output_pdf)
                zipf.writestr(filnavn, output_pdf.getvalue())
                
                # Neste målebrev starter på i
                startside = i

        # Håndter siste segment (fra startside til siste side)
        tekst_for_ett_malebrev = "\n".join(tekst_per_side[startside:len(tekst_per_side)])
        postnummer, mengde, dato = trekk_ut_verdier(tekst_for_ett_malebrev)
        filnavn = f"{postnummer}_{dato}.pdf"
        output_pdf = BytesIO()
        opprett_ny_pdf(combined_pdf, startside, len(tekst_per_side) - 1, output_pdf)
        zipf.writestr(filnavn, output_pdf.getvalue())

    zip_buffer.seek(0)
    return zip_buffer

# Streamlit-grensesnittet
st.title("Kombiner målebrev med vedlegg / Splitt og last ned som ZIP")

pdf_file = st.file_uploader("Last opp PDF-filen med Målebrev", type="pdf", key="combine_pdf")
folder_files = st.file_uploader("Last opp vedleggs-PDF-filer", type="pdf", accept_multiple_files=True, key="attachments")

if pdf_file and folder_files:
    # 1) Kombiner hoved-PDF og vedlegg
    combined_pdf = combine_pdf_and_attachments(pdf_file, folder_files)
    st.success("Kombinering fullført!")
    st.download_button(
        "Last ned kombinert PDF", 
        combined_pdf, 
        file_name="kombinert_dokument.pdf", 
        mime="application/pdf"
    )

    # 2) Splitt og last ned som ZIP
    if st.button("Start Splitting og Last ned som ZIP"):
        zip_buffer = split_pdf_to_zip(combined_pdf)
        st.success("Splitting fullført!")
        st.download_button(
            "Last ned alle splittede PDF-filer som ZIP", 
            zip_buffer, 
            file_name="Splittet_malebrev.zip", 
            mime="application/zip"
        )
