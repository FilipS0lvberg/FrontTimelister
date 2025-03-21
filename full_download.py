from create_prosjektregnsap import create_prosjektregnskap1
from create_timelister import create_timelister1
from write_erklering import write_erklering1
from supporting_functions import get_workers
import zipfile
import io


def download_all(year, selskapsnavn, prosjekttittel, prosjektnummer, prosjektleder, antall_ansatte, antall_arbeidspakker, prosjektbeskrivelse, godkjenningsperiode, prosjektstatusdato, prosjektstatus, godkjentdato, kategori, selskapvansker, kostnadsalternavier, erkleringdato):
    ######## Genererer timelister
    kvartaler = [1, 2, 3, 4]
    timelister = []
    for qtr in kvartaler:
        timeliste = create_timelister1(year, qtr, selskapsnavn, prosjekttittel, prosjektnummer, prosjektleder, get_workers(antall_ansatte), prosjektbeskrivelse, antall_arbeidspakker)
        timelister.append(timeliste)

    ##### Generere prosjektregnskap
    prosjektregnskap1 = create_prosjektregnskap1(year, selskapsnavn, prosjektnummer, prosjekttittel, godkjenningsperiode, prosjektstatusdato, prosjektstatus, godkjentdato, kategori, selskapvansker, kostnadsalternavier)

    ##### Generere erklering
    erklering1 = write_erklering1(selskapsnavn, erkleringdato, prosjektnummer, prosjektleder)

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for qtr, timeliste in enumerate(timelister):
            zipf.writestr(f"SF{year} - Timelister Q{qtr+1} - {selskapsnavn}.xlsx", timeliste)

        zipf.writestr(f'SF{year} - Prosjektregnskap - {selskapsnavn}.xlsx', prosjektregnskap1)
        zipf.writestr(f"SF{year} - Erklæring - {selskapsnavn}.docx", erklering1)

    # Get the zip file's bytes
    zip_buffer.seek(0)
    return zip_buffer.getvalue()

