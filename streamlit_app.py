import fitz  # PyMuPDF
import re
from datetime import datetime
from pathlib import Path
import zipfile
from io import BytesIO
import streamlit as st


# Hjelpefunksjoner
def les_tekst_fra_pdf(pdf_file):
    """Leser teksten fra hver side i PDF-filen."""
    pdf_file.seek(0)
    dokument = fitz.open(stream=pdf_file.read(), filetype="pdf")
    tekst_per_side = []
    for side_num in range(len(dokument)):
        side = dokument[side_num]
        tekst_per_side.append(side.get_text("text"))
    dokument.close()
    return tekst_per_side


def opprett_ny_pdf(original_pdf, startside, sluttside):
    """Lager en ny PDF med sidene fra start til slutt."""
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


def trekk_ut_verdier(tekst):
    """Henter postnummer, mengde og dato fra teksten."""
    postnummer_pattern = r'postnummer\s+beskrivelse\s+([\d.]+)'
    postnummer_match = re.search(postnummer_pattern, tekst, re.IGNORECASE)
    postnummer = postnummer_match.group(1).strip() if postnummer_match else "ukjent"

    mengde_pattern = r'Utført pr. d.d.:\s+([\d,.]+)'
    mengde_match = re.search(mengde_pattern, tekst)
    mengde = mengde_match.group(1) if mengde_match else "ukjent"

    dato_pattern = r'(\d{2}\.\d{2}\.\d{4})'
    dato_match = re.search(dato_pattern, tekst)
    dato = datetime.strptime(dato_match.group(1), "%d.%m.%Y").strftime("%Y%m%d") if dato_match else datetime.now().strftime("%Y%m%d")

    return postnummer, mengde, dato


def finn_vedlegg_i_tekst(tekst):
    """Finner vedleggsfilene fra 'Vedlagte dokumenter'."""
    vedlegg_mønster = r'Vedlagte dokumenter:\s+(.*?)($|\n\n)'
    match = re.search(vedlegg_mønster, tekst, re.IGNORECASE | re.DOTALL)
    if not match:
        return []
    vedlegg_tekst = match.group(1)
    pattern_pdf = r'([\w\-. ]+\.pdf)'
    return re.findall(pattern_pdf, vedlegg_tekst, re.IGNORECASE)


def split_malebrev_med_vedlegg(pdf_file, folder_files):
    """Splitt målebrev med alle sider og legg til vedlegg."""
    tekst_per_side = les_tekst_fra_pdf(pdf_file)
    pdf_file.seek(0)
    original_pdf = BytesIO(pdf_file.read())
    sidegrenser = []

    for i, tekst in enumerate(tekst_per_side):
        if "Målebrev" in tekst:
            sidegrenser.append(i)

    sidegrenser.append(len(tekst_per_side))

    folder_dict = {Path(file.name).name: file for file in folder_files}
    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for idx in range(len(sidegrenser) - 1):
            startside = sidegrenser[idx]
            sluttside = sidegrenser[idx + 1] - 1
            tekst_for_ett_malebrev = "\n".join(tekst_per_side[startside:sluttside + 1])

            postnummer, mengde, dato = trekk_ut_verdier(tekst_for_ett_malebrev)
            vedleggsliste = finn_vedlegg_i_tekst(tekst_for_ett_malebrev)

            pdf_bytes_malebrev = opprett_ny_pdf(original_pdf, startside, sluttside)
            combined_pdf = fitz.open(stream=pdf_bytes_malebrev.read(), filetype="pdf")

            for vedlegg in vedleggsliste:
                if vedlegg in folder_dict:
                    attachment = folder_dict[vedlegg]
                    attachment.seek(0)
                    vedlegg_pdf = fitz.open(stream=attachment.read(), filetype="pdf")
                    combined_pdf.insert_pdf(vedlegg_pdf)
                    vedlegg_pdf.close()

            final_pdf = BytesIO()
            combined_pdf.save(final_pdf)
            combined_pdf.close()
            final_pdf.seek(0)

            filnavn = f"{postnummer}_{dato}.pdf"
            zipf.writestr(filnavn, final_pdf.getvalue())

    zip_buffer.seek(0)
    return zip_buffer


# Streamlit-grensesnitt
st.title("Splitt målebrev med vedlegg")

pdf_file = st.file_uploader("Last opp PDF-filen med målebrev", type="pdf")
vedlegg_files = st.file_uploader("Last opp vedlegg", type="pdf", accept_multiple_files=True)

if pdf_file and vedlegg_files:
    if st.button("Splitt og last ned ZIP"):
        zip_buffer = split_malebrev_med_vedlegg(pdf_file, vedlegg_files)
        st.success("Splitting fullført!")
        st.download_button("Last ned ZIP", zip_buffer, "splittet_malebrev.zip", "application/zip")
