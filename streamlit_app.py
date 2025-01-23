import fitz  # PyMuPDF
import re
from datetime import datetime
from pathlib import Path
import zipfile
import streamlit as st
from io import BytesIO

# ----------- 1) Hjelpefunksjoner for parsing av tekst -----------------

def trekk_ut_verdier(tekst):
    """
    Trekk ut postnummer, mengde, dato fra en (sammenhengende) streng.
    Tilpass regex om PDF-en har litt annen struktur.
    """
    postnummer_pattern = r'postnummer\s+beskrivelse\s+([\d.]+)\s+(?=rs|stk|kg|m|m2|m3)\b'
    postnummer_match = re.search(postnummer_pattern, tekst, re.IGNORECASE)
    postnummer = postnummer_match.group(1).strip() if postnummer_match else "ukjent"

    mengde_pattern = r'(?<=Utført pr. d.d.:\n)([\d,]+)'
    mengde_match = re.search(mengde_pattern, tekst)
    mengde = mengde_match.group(1) if mengde_match else "ukjent"

    dato_pattern = r'(\d{2}\.\d{2}\.\d{4})'
    dato_match_str = datetime.now().strftime("%Y%m%d")
    # Dersom en faktisk dato finnes i teksten, overskriv default
    if (dato_re_match := re.search(dato_pattern, tekst)):
        dato_match_str = datetime.strptime(dato_re_match.group(1), "%d.%m.%Y").strftime("%Y%m%d")

    return postnummer, mengde, dato_match_str


def finn_vedlegg_i_tekst(tekst):
    """
    Finn alle PDF-vedlegg som listes etter 'Vedlagte dokumenter:'.
    Returnerer en liste med filnavn (f.eks. ["test1.pdf", "oversikt.pdf", ...]).
    
    Hvis du har flere 'Vedlagte dokumenter:'-seksjoner i samme målebrev,
    kan du justere til re.findall for å plukke ut alle, eller løkke over.
    """
    vedlegg_mønster = r'Vedlagte dokumenter:(.*?)(?=\n\n|$)'  # Litt forsiktig
    match = re.search(vedlegg_mønster, tekst, re.IGNORECASE|re.DOTALL)
    if not match:
        return []

    vedlegg_tekst = match.group(1)

    # Finn alt som ender på ".pdf", evt. med bindestrek, mellomrom i navnet, etc.
    pattern_pdf = r'([\w\-. ]+\.pdf)'
    found = re.findall(pattern_pdf, vedlegg_tekst, re.IGNORECASE)

    found_cleaned = []
    for f in found:
        # Rens filnavn (f.eks. "c:/mappe/vedlegg.pdf" -> "vedlegg.pdf")
        fname = f.strip().replace("\\", "/").split("/")[-1]
        found_cleaned.append(fname)

    return found_cleaned


def les_tekst_fra_pdf(pdf_file):
    """
    Leser all tekst side for side i en PDF (lastet opp via Streamlit).
    Returnerer en liste 'tekst_per_side', der indeks 0 = side 0, 1 = side 1, ...
    """
    pdf_file.seek(0)
    dokument = fitz.open(stream=pdf_file.read(), filetype="pdf")
    tekst_per_side = []
    for side_num in range(len(dokument)):
        side = dokument[side_num]
        tekst_per_side.append(side.get_text("text"))
    dokument.close()
    return tekst_per_side

# ----------- 2) Hjelpefunksjon for å lage en PDF av utvalgte sider  -----------------

def opprett_ny_pdf(original_pdf, startside, sluttside):
    """
    Returnerer en BytesIO-PDF med sidene [startside .. sluttside] (inkludert).
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

# ----------- 3) Funksjon som splitter PDF-en pr. målebrev + vedlegg  -----------------

def split_malebrev_med_vedlegg(pdf_file, folder_files):
    """
    - Leter etter "Målebrev" pr side for å finne start på nye målebrev
      (side i -> side i+1 -> ... til vi møter neste "Målebrev" eller slutt).
    - For hver bolk (startside -> sluttside):
        1) Samle all tekst
        2) Finn postnr, mengde, dato
        3) Finn vedleggsfilnavn fra "Vedlagte dokumenter:"
        4) Lag en PDF med de sidene
        5) Append vedleggene bak
        6) Legg resultatet i en ZIP
    - Returnerer en BytesIO som inneholder ZIP-en.
    """

    # 1) Les teksten pr side -> en liste
    tekst_per_side = les_tekst_fra_pdf(pdf_file)

    # 2) Les inn PDF-dataen i minnet (slik at vi kan hente ut sider flere ganger)
    pdf_file.seek(0)
    original_pdf_data = pdf_file.read()
    original_pdf_mem = BytesIO(original_pdf_data)

    # 3) Finn sidegrenser: Hver gang vi ser "Målebrev" i en side, regner vi det som start på ny post
    sidegrenser = []
    for i, tekst in enumerate(tekst_per_side):
        # Obs: "Målebrev" kan forekomme midt i siden, men da bruker vi "if 'Målebrev' in tekst"
        # Hvis det er store bokstaver, bruk "MÅLEBREV" -> da juster du til re.IGNORECASE
        if "Målebrev" in tekst:
            sidegrenser.append(i)

    if not sidegrenser:
        # Ingen "Målebrev" funnet = ingen splitting
        # Returner tom ZIP eller handle det på en annen måte
        empty_zip = BytesIO()
        with zipfile.ZipFile(empty_zip, 'w') as zf:
            pass
        empty_zip.seek(0)
        return empty_zip

    # Legg til en "sluttgrense" på slutten
    sidegrenser.append(len(tekst_per_side))

    # 4) Organiser filer i en dictionary: {filnavn.pdf: fileobj}
    folder_dict = {Path(file.name).name: file for file in folder_files}

    # 5) Opprett en ZIP i minnet
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:

        # Loop over målebrev
        for idx in range(len(sidegrenser) - 1):
            startside = sidegrenser[idx]
            sluttside = sidegrenser[idx+1] - 1

            # 5.1) Samle all tekst for dette målebrevet
            tekst_for_ett_malebrev = "\n".join(tekst_per_side[startside:sluttside+1])

            # 5.2) Finn postnummer, mengde, dato
            postnummer, mengde, dato = trekk_ut_verdier(tekst_for_ett_malebrev)

            # 5.3) Finn hvilke vedlegg (pdf-filnavn) som står under "Vedlagte dokumenter:"
            vedleggsliste = finn_vedlegg_i_tekst(tekst_for_ett_malebrev)

            # 5.4) Lag PDF med disse sidene
            pdf_bytes_malebrev = opprett_ny_pdf(original_pdf_mem, startside, sluttside)

            # 5.5) Åpne den nye PDF-en, append vedlegg
            combined = fitz.open(stream=pdf_bytes_malebrev.read(), filetype="pdf")

            for filnavn in vedleggsliste:
                if filnavn in folder_dict:
                    attach_file = folder_dict[filnavn]
                    attach_file.seek(0)
                    attach_pdf = fitz.open(stream=attach_file.read(), filetype="pdf")
                    combined.insert_pdf(attach_pdf)
                    attach_pdf.close()

            # 5.6) Lagre finalen i minnet
            final_pdf = BytesIO()
            combined.save(final_pdf)
            combined.close()
            final_pdf.seek(0)

            # 5.7) Legg i ZIP, f.eks. filnavn: "<postnr>_<dato>.pdf"
            zipf.writestr(f"{postnummer}_{dato}.pdf", final_pdf.getvalue())

    zip_buffer.seek(0)
    return zip_buffer

# ----------- 4) Streamlit-grensesnitt  -----------------

st.title("Splitt fler-siders målebrev + vedlegg pr post")

pdf_file = st.file_uploader("Last opp PDF-filen med målebrev", type="pdf")
vedlegg_files = st.file_uploader("Last opp vedlegg (PDF)", type="pdf", accept_multiple_files=True)

if pdf_file and vedlegg_files:

    # (Valgfritt) en debug-knapp for å se tekst pr side
    if st.button("Debug: Vis tekst pr side"):
        test_tekst = les_tekst_fra_pdf(pdf_file)
        for i, side_tekst in enumerate(test_tekst):
            st.write(f"--- Side {i} ---")
            st.write(side_tekst)
        st.write("Scroll opp/ned for å se hva som faktisk står på hver side.")

    if st.button("Splitt og Last ned ZIP"):
        zip_buffer = split_malebrev_med_vedlegg(pdf_file, vedlegg_files)
        st.success("Splitting ferdig!")
        st.download_button(
            label="Last ned ZIP",
            data=zip_buffer,
            file_name="splittet_malebrev.zip",
            mime="application/zip",
        )
