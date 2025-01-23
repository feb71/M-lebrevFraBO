import fitz  # PyMuPDF
import re
from datetime import datetime
from pathlib import Path
import zipfile
from io import BytesIO
import streamlit as st

# ----------- Hjelpefunksjoner -----------

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


def trekk_ut_vedlegg(tekst):
    """
    Finner alle vedlegg som står under "Vedlagte dokumenter".
    Søker i hele teksten (samlet fra alle sider i målebrevet).
    """
    vedlegg_mønster = r"Vedlagte dokumenter:(.*?)(Skrevet ut|Signatur|$)"
    match = re.search(vedlegg_mønster, tekst, re.IGNORECASE | re.DOTALL)
    if not match:
        return []

    vedlegg_tekst = match.group(1)
    # Finner alle filnavn som ender med ".pdf", uavhengig av linjeskift
    pattern_pdf = r"([\w\-. ]+\.pdf)"
    return re.findall(pattern_pdf, vedlegg_tekst, re.IGNORECASE)


def split_malebrev_med_vedlegg(pdf_file, folder_files):
    """
    Splitt målebrev basert på "Målebrev" og legg ved alle relevante vedlegg.
    Håndterer flere sider i ett målebrev.
    """
    tekst_per_side = les_tekst_fra_pdf(pdf_file)
    pdf_file.seek(0)
    original_pdf = BytesIO(pdf_file.read())

    # Finn grenser for målebrev basert på "Målebrev" i teksten
    sidegrenser = []
    for i, tekst in enumerate(tekst_per_side):
        if "Målebrev" in tekst:
            sidegrenser.append(i)

    # Legg til siste side som sluttgrense
    sidegrenser.append(len(tekst_per_side))

    folder_dict = {Path(file.name).name: file for file in folder_files}
    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for idx in range(len(sidegrenser) - 1):
            startside = sidegrenser[idx]
            sluttside = sidegrenser[idx + 1] - 1

            # Kombiner tekst fra alle sider i målebrevet
            tekst_for_malebrev = "\n".join(tekst_per_side[startside:sluttside + 1])

            # Hent vedlegg fra teksten
            vedleggsliste = trekk_ut_vedlegg(tekst_for_malebrev)

            # Lag en ny PDF med alle sidene i dette målebrevet
            pdf_bytes_malebrev = opprett_ny_pdf(original_pdf, startside, sluttside)
            combined_pdf = fitz.open(stream=pdf_bytes_malebrev.read(), filetype="pdf")

            # Legg til vedleggene i PDF-en
            for vedlegg in vedleggsliste:
                if vedlegg in folder_dict:
                    attachment = folder_dict[vedlegg]
                    attachment.seek(0)
                    vedlegg_pdf = fitz.open(stream=attachment.read(), filetype="pdf")
                    combined_pdf.insert_pdf(vedlegg_pdf)
                    vedlegg_pdf.close()

            # Lagre PDF for dette målebrevet
            final_pdf = BytesIO()
            combined_pdf.save(final_pdf)
            combined_pdf.close()
            final_pdf.seek(0)

            # Legg PDF-en i ZIP-filen
            filnavn = f"Malebrev_{startside + 1}_til_{sluttside + 1}.pdf"
            zipf.writestr(filnavn, final_pdf.getvalue())

    zip_buffer.seek(0)
    return zip_buffer

# ----------- Streamlit-grensesnitt -----------

st.title("Splitt målebrev med vedlegg")

pdf_file = st.file_uploader("Last opp PDF-filen med målebrev", type="pdf")
vedlegg_files = st.file_uploader("Last opp vedlegg", type="pdf", accept_multiple_files=True)

if pdf_file and vedlegg_files:
    # Debug: Vis tekst per side
    if st.button("Vis tekst per side"):
        tekst_per_side = les_tekst_fra_pdf(pdf_file)
        for i, tekst in enumerate(tekst_per_side):
            st.write(f"--- Side {i + 1} ---")
            st.write(tekst)

    if st.button("Splitt og last ned ZIP"):
        zip_buffer = split_malebrev_med_vedlegg(pdf_file, vedlegg_files)
        st.success("Splitting fullført!")
        st.download_button(
            label="Last ned ZIP",
            data=zip_buffer,
            file_name="splittet_malebrev.zip",
            mime="application/zip",
        )
