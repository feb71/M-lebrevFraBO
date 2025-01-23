import fitz  # PyMuPDF
import re
from datetime import datetime
from pathlib import Path
import os
import zipfile
import streamlit as st
from io import BytesIO

# --------------------------------------------------------------------------------
# 1) Funksjon for å hente ut info (postnummer, mengde, dato) fra en samlet streng
# --------------------------------------------------------------------------------
def trekk_ut_verdier(tekst):
    postnummer_pattern = r'postnummer\s+beskrivelse\s+([\d.]+)\s+(?=rs|stk|kg|m|m2|m3)\b'
    postnummer_match = re.search(postnummer_pattern, tekst, re.IGNORECASE)
    postnummer = postnummer_match.group(1).strip() if postnummer_match else "ukjent"

    mengde_pattern = r'(?<=Utført pr. d.d.:\n)([\d,]+)'
    mengde_match = re.search(mengde_pattern, tekst)
    mengde = mengde_match.group(1) if mengde_match else "ukjent"

    # Dato: bruk dagens dato (YYYYMMDD) hvis ingen dato i teksten
    dato_pattern = r'(\d{2}\.\d{2}\.\d{4})'
    dato_match_str = datetime.now().strftime("%Y%m%d")
    if (dato_re_match := re.search(dato_pattern, tekst)):
        dato_match_str = datetime.strptime(dato_re_match.group(1), "%d.%m.%Y").strftime("%Y%m%d")

    return postnummer, mengde, dato_match_str

# --------------------------------------------------------------------------------
# 2) Funksjon for å hente ut PDF-vedleggene som er listet i en streng med "Vedlagte dokumenter:"
# --------------------------------------------------------------------------------
def finn_vedlegg_i_tekst(tekst):
    """
    Ser etter 'Vedlagte dokumenter:' og henter ut alle *.pdf i samme avsnitt.
    Returnerer en liste med filnavn, f.eks. ["vedlegg1.pdf", "rapport.pdf", ...].
    Juster regex/fremgangsmåte ved behov.
    """
    # Eksempel: Du kan finne alt som står etter "Vedlagte dokumenter:" på samme side.
    # Hvis du har flere "Vedlagte dokumenter:"-seksjoner per målebrev, må du iterere.
    vedlegg_mønster = r'Vedlagte dokumenter:(.*?)(?=\n\n|$)'  # litt forsiktig
    match = re.search(vedlegg_mønster, tekst, re.IGNORECASE|re.DOTALL)
    if not match:
        return []

    # Teksten i gruppa kan inneholde flere linjer, f.eks.:
    # vedlegg1.pdf
    # vedlegg2.pdf
    # ...
    # Splitt i linjer og hent ut .pdf
    vedlegg_tekst = match.group(1)

    # Finn alt som ender på ".pdf" (enten store eller små bokstaver)
    # og fjern eventuelle path-biter før.
    pattern_pdf = r'([\w\-. ]+\.pdf)'
    found = re.findall(pattern_pdf, vedlegg_tekst, re.IGNORECASE)
    
    # Rens ut spacing, fjerne "\\"-stier og lignende
    found_cleaned = []
    for f in found:
        # Ta kun selve filnavnet (f.eks. filnavn.pdf, selv om det sto c:/mappe/filnavn.pdf)
        # Eksempel: "some/folder/test.pdf" -> "test.pdf"
        fname = f.strip().replace("\\", "/").split("/")[-1]
        found_cleaned.append(fname)

    return found_cleaned

# --------------------------------------------------------------------------------
# 3) Les all tekst fra PDF pr side
# --------------------------------------------------------------------------------
def les_tekst_fra_pdf(pdf_file):
    pdf_file.seek(0)
    dokument = fitz.open(stream=pdf_file.read(), filetype="pdf")
    tekst_per_side = []
    for side_num in range(len(dokument)):
        side = dokument[side_num]
        tekst_per_side.append(side.get_text("text"))
    dokument.close()
    return tekst_per_side

# --------------------------------------------------------------------------------
# 4) Opprett ny PDF fra utvalgte sider (startside -> sluttside)
# --------------------------------------------------------------------------------
def opprett_ny_pdf(original_pdf, startside, sluttside):
    """
    Returnerer en BytesIO som inneholder PDF-sidene fra startside til sluttside (inkludert).
    """
    output_pdf = BytesIO()
    original_pdf.seek(0)
    dokument = fitz.open(stream=original_pdf.read(), filetype="pdf")
    ny_pdf = fitz.open()
    ny_pdf.insert_pdf(dokument, from_page=startside, to_page=sluttside)
    ny_pdf.save(output_pdf)
    ny_pdf.close()
    dokument.close()
    output_pdf.seek(0)
    return output_pdf

# --------------------------------------------------------------------------------
# 5) Hoved-funksjon som splitter PDF-en pr målebrev, og for hvert målebrev
#    legger vi på de vedleggene som er listet under "Vedlagte dokumenter:" 
#    (uansett hvilken side av målebrevet de sto på).
# --------------------------------------------------------------------------------
def split_malebrev_med_vedlegg(pdf_file, folder_files):
    """
    - Finne grenser for hvert målebrev (basert på ordet "Målebrev" i PDF-en).
    - For hver post: 
       a) samle ALLE sidene (startside -> sluttside),
       b) parse ut "Vedlagte dokumenter" fra all teksten i målebrevet,
       c) appender de vedlagte PDF-filene, 
       d) navngir filen f"{postnummer}_{dato}.pdf".
    - Returnerer en ZIP i en BytesIO.
    """
    # 1) Les inn PDF-tekst pr. side
    tekst_per_side = les_tekst_fra_pdf(pdf_file)

    # For å kunne lage nye PDF-er i prosessen flere ganger, 
    # må vi ha "original_pdf" i minnet som BytesIO
    pdf_file.seek(0)
    original_pdf_data = pdf_file.read()
    original_pdf_bytesio = BytesIO(original_pdf_data)

    # 2) Finn sidegrenser for målebrev
    sidegrenser = []
    for i, tekst in enumerate(tekst_per_side):
        if "Målebrev" in tekst:
            sidegrenser.append(i)

    # Hvis PDF-en starter med "Målebrev" på side 0, men vi trenger en "sluttgrense" i tillegg
    # Ved splitting er logikken: [sidegrenser[0], sidegrenser[1]-1], [sidegrenser[1], sidegrenser[2]-1], ...
    # men siste målebrev strekker seg til siste side i dokumentet
    # sidegrenser = [0, 3, 7, ...]
    # Sjekk at vi har minst én post
    if not sidegrenser:
        # Fant ingen "Målebrev", da kan man håndtere det som man vil. 
        # Returnerer tom ZIP i dette eksemplet.
        empty_zip = BytesIO()
        with zipfile.ZipFile(empty_zip, "w", zipfile.ZIP_DEFLATED) as zf:
            pass
        empty_zip.seek(0)
        return empty_zip

    # Legg til en "kunstig" sluttgrense på siste side + 1
    sidegrenser.append(len(tekst_per_side))

    # 3) Opprett en ZIP i minnet og fyll den med PDF-er
    zip_buffer = BytesIO()
    folder_dict = {Path(file.name).name: file for file in folder_files}

    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        # Gå gjennom par av grenser: (sidegrenser[i], sidegrenser[i+1]) -> definisjon av et målebrev
        for idx in range(len(sidegrenser) - 1):
            startside = sidegrenser[idx]
            sluttside = sidegrenser[idx+1] - 1

            # Samle tekst fra alle sidene i dette målebrevet
            tekst_for_ett_malebrev = "\n".join(tekst_per_side[startside:sluttside+1])

            # Trekk ut info om postnummer, dato, mengde
            postnummer, mengde, dato = trekk_ut_verdier(tekst_for_ett_malebrev)

            # Finn vedlegg (alle .pdf-filer listet under "Vedlagte dokumenter:")
            vedlegg_fra_tekst = finn_vedlegg_i_tekst(tekst_for_ett_malebrev)

            # Lag en PDF med sidene som hører til målebrevet
            malebrev_pdf_bytes = opprett_ny_pdf(original_pdf_bytesio, startside, sluttside)
            malebrev_pdf = fitz.open(stream=malebrev_pdf_bytesio.read(), filetype="pdf")
            malebrev_pdf.close()

            # Nå skal vi opprette en "endelig" PDF i minnet der vi først 
            # legger inn målebrevsidene, deretter vedleggene. 
            combined_for_this_post = fitz.open(stream=malebrev_pdf_bytes, filetype="pdf")

            # For hvert vedleggsnavn -> finn i folder_dict og append
            for filnavn in vedlegg_fra_tekst:
                if filnavn in folder_dict:
                    attachment_file = folder_dict[filnavn]
                    attachment_file.seek(0)
                    attachment_pdf = fitz.open(stream=attachment_file.read(), filetype="pdf")

                    # Append attachment til combined
                    combined_for_this_post.insert_pdf(attachment_pdf)
                    attachment_pdf.close()

            # Lag en bytesIO av combined
            final_pdf_for_post = BytesIO()
            combined_for_this_post.save(final_pdf_for_post)
            combined_for_this_post.close()
            final_pdf_for_post.seek(0)

            # Sett filnavn for ZIP
            filnavn = f"{postnummer}_{dato}.pdf"
            zipf.writestr(filnavn, final_pdf_for_post.getvalue())

    zip_buffer.seek(0)
    return zip_buffer

# --------------------------------------------------------------------------------
# 6) Streamlit-grensesnittet
# --------------------------------------------------------------------------------
st.title("Målebrev-splitting med vedlegg per post")

pdf_file = st.file_uploader("Last opp PDF-filen med Målebrev", type="pdf")
folder_files = st.file_uploader("Last opp vedlegg (PDF)", type="pdf", accept_multiple_files=True)

if pdf_file and folder_files:
    if st.button("Splitt og last ned ZIP"):
        # Kjør funksjonen som lager ZIP
        zip_buffer = split_malebrev_med_vedlegg(pdf_file, folder_files)
        st.success("Splitting fullført!")
        st.download_button(
            "Last ned ZIP med splittede målebrev + vedlegg",
            zip_buffer,
            file_name="Splittet_malebrev.zip",
            mime="application/zip"
        )
