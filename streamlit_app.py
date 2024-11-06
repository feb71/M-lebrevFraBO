import re
from datetime import datetime
import fitz  # PyMuPDF
import os
import zipfile
import streamlit as st

st.set_page_config(layout="wide")  # Bruk hele bredden av skjermen

# Funksjon for å trekke ut verdier fra teksten
def trekk_ut_verdier(tekst):
    # Regex for å finne postnummeret basert på "Postnummer Beskrivelse" og enhetstype
    postnummer_pattern = r'postnummer\s+beskrivelse\s+([\d.]+)\s+(?=rs|stk|kg|m|m2|m3)\b'
    postnummer_match = re.search(postnummer_pattern, tekst, re.IGNORECASE)

    postnummer = postnummer_match.group(1).strip() if postnummer_match else "ukjent"

    # Logg funnet postnummer for å bekrefte
    st.write(f"Funnet postnummer: {postnummer}")

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

# Hovedlogikk for å behandle PDF-splitting
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

# Funksjon for å lese tekst fra PDF for splitting
def les_tekst_fra_pdf(pdf_file):
    dokument = fitz.open(stream=pdf_file.read(), filetype="pdf")
    tekst_per_side = []
    for side_num in range(len(dokument)):
        side = dokument[side_num]
        tekst_per_side.append(side.get_text("text"))
    dokument.close()
    return tekst_per_side

# Streamlit-app for å laste opp PDF og velge lagringssted
st.title("Splitt PDF-fil pr post")
uploaded_pdf = st.file_uploader("Last opp PDF-fil for splitting", type="pdf")
output_folder = os.path.join(os.path.expanduser("~"), "Downloads", "Splittet_malebrev")

if uploaded_pdf and st.button("Start Splitting av PDF"):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    st.write(f"Starter splitting og lagrer i mappe: {output_folder}")
    opprettede_filer = behandle_og_splitte_pdf(uploaded_pdf, output_folder)

    st.success("Splitting fullført!")
    st.write(f"Antall opprettede filer: {len(opprettede_filer)}")

    # Gi brukeren mulighet til å laste ned ZIP-fil
    zip_path = os.path.join(os.path.expanduser("~"), "Downloads", "Splittet_malebrev.zip")
    try:
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for fil in opprettede_filer:
                zipf.write(fil, os.path.basename(fil))
    except Exception as e:
        st.error(f"En feil oppstod under opprettelse av ZIP-fil: {e}")

    if os.path.exists(zip_path):
        with open(zip_path, "rb") as z:
            st.download_button("Last ned alle PDF-filer som ZIP", z, file_name="Splittet_malebrev.zip", mime="application/zip")
    else:
        st.error("ZIP-filen ble ikke opprettet.")
