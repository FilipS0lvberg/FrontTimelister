import streamlit as st
from supporting_functions import get_workers
from create_timelister import create_timelister1
from create_prosjektregnsap import create_prosjektregnskap1
from write_erklering import write_erklering1
from full_download import download_all


st.title("Front Generation - Rapportering :memo:")
########### INITIALIZE ALL SESSION STATE VARIABLES ###########

# Initialize the flag in session state if it doesn't exist
if "flag" not in st.session_state:
    st.session_state.flag = False

if "year" not in st.session_state:
    st.session_state.year = 2025

if "selskapsnavn" not in st.session_state:
    st.session_state.selskapsnavn = ""

if "prosjekttittel" not in st.session_state:
    st.session_state.prosjekttittel = ""

if "prosjektnummer" not in st.session_state:
    st.session_state.prosjektnummer = ""

if "prosjektleder" not in st.session_state:
    st.session_state.prosjektleder = ""

if "antall_ansatte" not in st.session_state:
    st.session_state.antall_ansatte = 1

if "antall_arbeidspakker" not in st.session_state:
    st.session_state.antall_arbeidspakker = 1

if "prosjektbeskrivelse" not in st.session_state:
    st.session_state.prosjektbeskrivelse = ""

if "godkjenningsperiode" not in st.session_state:
    st.session_state.godkjenningsperiode = ""

if "prosjektstatusdato" not in st.session_state:
    st.session_state.prosjektstatusdato = ""

if "prosjektstatus" not in st.session_state:
    st.session_state.prosjektstatus = "Godkjent"

if "godkjentdato" not in st.session_state:
    st.session_state.godkjentdato = ""

if "kategori" not in st.session_state:
    st.session_state.kategori = ""

if "selskapvansker" not in st.session_state:
    st.session_state.selskapvansker = "Nei"

if "kostnadsalternativer" not in st.session_state:
    st.session_state.kostnadsalternativer = ""

if "erkleringdato" not in st.session_state:
    st.session_state.erkleringdato = ""



tab1, tab2, tab3, tab4 = st.tabs(["Timelister", "Prosjektregnskap", "Erklæring", "Full nedlastning"]) # DEFINE THE TABS

# TIMELISTER TAB
with tab1:
    # FRONTEND INPUT VARIABLES FOR GENERATING TIMELISTER 
    year = st.number_input("Årstall", min_value=2020, max_value=2050, value=st.session_state.year)
    selskapsnavn = st.text_input("Selskapsnavn")
    st.session_state.selskapsnavn = selskapsnavn

    prosjekttittel = st.text_input("Prosjekttittel")
    st.session_state.prosjekttittel = prosjekttittel

    prosjektleder = st.text_input("Prosjektleder")
    st.session_state.prosjektleder = prosjektleder

    prosjektnummer = st.number_input("Prosjektnummer", value=1, min_value=1)
    st.session_state.prosjektnummer = prosjektnummer

    ant_ansatte = st.number_input("Antall ansatte", value=1, min_value=1)
    st.session_state.antall_ansatte = ant_ansatte

    antall_arbeidspakker = st.number_input("Antall Arbeidspakker", value=1, min_value=1)
    st.session_state.antall_arbeidspakker = antall_arbeidspakker

    prosjektbeskrivelse = st.text_area("Prosjektbeskrivelse", height=700, help="Denne boksen skal inneholde prosjektets hovedmål, samt tittel på arbeidspakker og aktiviteter")
    st.session_state.prosjektbeskrivelse = prosjektbeskrivelse


    if st.session_state.flag == False:
        autentisering = st.text_input("Passord", type="password")
        if st.button("Autentiser"):
            if autentisering == "Front1234":
                st.session_state.flag = True
                st.rerun()
            

    if st.session_state.flag:
        ####################### DOWNLOAD BUTTON FOR DOWNLOADING WORD DOC
        st.download_button(
            type="primary",
            label="Last ned Q1 - Timelister",
            # data=write_docx(st.session_state.norsk_proj_title, st.session_state.main_checkboxes ,st.session_state.main_box_contents, st.session_state.workpackages),
            data=create_timelister1(year=year, quarter=1, selskap=selskapsnavn, prosjekttittel=prosjekttittel, prosjektnummer=prosjektnummer, prosjektleder=prosjektleder, prosject_description=prosjektbeskrivelse, workers=get_workers(ant_ansatte), num_workpackages=antall_arbeidspakker),
            file_name=f"SF{year} - Timelister Q1 - {selskapsnavn}.xlsx",
            mime="xlsx",
            key="download0"
        )

        st.download_button(
            type="primary",
            label="Last ned Q2 - Timelister",
            # data=write_docx(st.session_state.norsk_proj_title, st.session_state.main_checkboxes ,st.session_state.main_box_contents, st.session_state.workpackages),
            data=create_timelister1(year=year, quarter=2, selskap=selskapsnavn, prosjekttittel=prosjekttittel, prosjektnummer=prosjektnummer, prosjektleder=prosjektleder, prosject_description=prosjektbeskrivelse, workers=get_workers(ant_ansatte), num_workpackages=antall_arbeidspakker),
            file_name=f"SF{year} - Timelister Q2 - {selskapsnavn}.xlsx",
            mime="xlsx",
            key="download1"
        )

        st.download_button(
            type="primary",
            label="Last ned Q3 - Timelister",
            # data=write_docx(st.session_state.norsk_proj_title, st.session_state.main_checkboxes ,st.session_state.main_box_contents, st.session_state.workpackages),
            data=create_timelister1(year=year, quarter=3, selskap=selskapsnavn, prosjekttittel=prosjekttittel, prosjektnummer=prosjektnummer, prosjektleder=prosjektleder, prosject_description=prosjektbeskrivelse, workers=get_workers(ant_ansatte), num_workpackages=antall_arbeidspakker),
            file_name=f"SF{year} - Timelister Q3 - {selskapsnavn}.xlsx",
            mime="xlsx",
            key="download2"
        )

        st.download_button(
            type="primary",
            label="Last ned Q4 - Timelister",
            # data=write_docx(st.session_state.norsk_proj_title, st.session_state.main_checkboxes ,st.session_state.main_box_contents, st.session_state.workpackages),
            data=create_timelister1(year=year, quarter=4, selskap=selskapsnavn, prosjekttittel=prosjekttittel, prosjektnummer=prosjektnummer, prosjektleder=prosjektleder, prosject_description=prosjektbeskrivelse, workers=get_workers(ant_ansatte), num_workpackages=antall_arbeidspakker),
            file_name=f"SF{year} - Timelister Q4 - {selskapsnavn}.xlsx",
            mime="xlsx",
            key="download3"
        )

# PROSJEKTREGNSKAP TAB
with tab2:
    # FRONTEND INPUT VARIABLES FOR GENERATING PROSJEKTREGNSKAP
    year1 = st.number_input("Årstall", min_value=2020, max_value=2050, value=st.session_state.year, help="Det året du vil at prosjektregnskapet skal gjelde for", key="year1")
    selskapsnavn1 = st.text_input("Juridisk selskapsnavn", value=st.session_state.selskapsnavn, key="selskapsnavn1", help="Husk å inkludere AS/ASA")
    prosjekttittel1 = st.text_input("Prosjekttittel", value=st.session_state.prosjekttittel, key="prosjekttittel1")
    prosjektnummer1 = st.number_input("Prosjektnummer", value=st.session_state.prosjektnummer, min_value=1, key="prosjektnummer1")

    godkjenningsperiode = st.text_input("Godkjennelsesperiode (dato fra - dato til)")
    st.session_state.godkjenningsperiode = godkjenningsperiode

    prosjektstatusdato = st.text_input("Prosjektstatusdato", help="Dato for utarbeidelse av prosjektregnskap")
    st.session_state.prosjektstatusdato = prosjektstatusdato

    prosjektstatus = "Godkjent"

    godkjentdato = st.text_input("Søknad godkjent dato", help="Dato for godkjennelse av søknad, er oppført i godkjenningsbrevet fra NFR")
    st.session_state.godkjentdato = godkjentdato

    kategori = st.radio("Prosjektkategori", ["Utviklingsprosjekt eksperimentell utvikling", "Innovasjonsprosjekt industriell forskning"])
    st.session_state.kategori = kategori

    selskapvansker = st.radio("Var selskapet i vansker ved godkjenningstidspunkt målt etter siste avlagte årsregnskap?", ["Ja", "Nei"], index=1, help="Hvis ja, legg ved dokumentasjon på at selskapet ikke er i vansker")
    st.session_state.selskapvansker = selskapvansker
    
    kostnads_alternativer = ["Timekostnader", "Prosjektkostnader", "Bruk av eget utstyr", "Kapitalkostnader"]
    kost_opts = st.pills("Kostnadsfaner", kostnads_alternativer, selection_mode="multi")
    st.session_state.kostnadsalternativer = kost_opts


    # if st.button("Sjekk sessionstate"):
    #     st.write(st.session_state.prosjekttittel)

    if st.session_state.flag == False:
        autentisering = st.text_input("Passord", type="password", key="prosjektregnskappassord")
        if st.button("Autentiser", key="prosjektregnskapautentiser"):
            if autentisering == "Front1234":
                st.session_state.flag = True
                st.rerun()

    if st.session_state.flag:
        st.download_button(
            type="primary",
            label=f"Last ned {year} Prosjektregnskap",
            data=create_prosjektregnskap1(year=year1, 
                                          juridisknavn=selskapsnavn1,
                                          project_number=prosjektnummer1, 
                                          project_title=prosjekttittel1,
                                          godkjent_periode=godkjenningsperiode,
                                          status_dato=prosjektstatusdato,
                                          status=prosjektstatus,
                                          godkjent_dato=godkjentdato,
                                          kategori=kategori,
                                          selskap_vansker=selskapvansker,
                                          cost_options=kost_opts),
            file_name=f"SF{year} - Prosjektregnskap - {st.session_state.selskapsnavn}.xlsx",
            key="download4",
            mime="xlsx"
        )


# ERKLERING TAB
with tab3:
    # FRONTEND INPUT VARIABLES FOR GENERATING ERKLERING
    year2 = st.number_input("Årstall", min_value=2020, max_value=2050, value=st.session_state.year, help="Det året du vil at prosjektregnskapet skal gjelde for", key="year2")
    
    erkleringdato = st.text_input("Dato for erklæring")
    st.session_state.erkleringdato = erkleringdato
    
    selskapsnavn2 = st.text_input("Juridisk selskapsnavn", value=st.session_state.selskapsnavn, key="selskapsnavn2")
    prosjektnummer2 = st.number_input("Prosjektnummer", value=st.session_state.prosjektnummer, min_value=1, key="prosjektnummer2")
    prosjektleder2 = st.text_input("Prosjektleder", value=st.session_state.prosjektleder, key="prosjektlder2")
    
    if st.session_state.flag == False:
        autentisering = st.text_input("Passord", type="password", key="erkleringpassord")
        if st.button("Autentiser", key="erkleringautoriser"):
            if autentisering == "Front1234":
                st.session_state.flag = True
                st.rerun()

    if st.session_state.flag:
        st.download_button(
            type="primary",
            label=f"Last ned {year} Erklæring",
            data=write_erklering1(dato=erkleringdato, selskapsnavn=selskapsnavn2, projectnumber=prosjektnummer2, prosjektleder=prosjektleder2),
            file_name=f"SF{year2} - Erklæring - {st.session_state.selskapsnavn}.docx",
            mime="docx",
            key="download5"
        )


# DOWNLOAD ALL TAB (KOMPRIMERT MAPPE .ZIP)
with tab4:
    st.write("Last ned komprimert mappe med alle filer :rocket:")

    if st.session_state.flag == False:
        autentisering = st.text_input("Passord", type="password", key="downallpassword")  
        if st.button("Autentiser", key="downallautoriser"):
            if autentisering == "Front1234":
                st.session_state.flag = True
                st.rerun()

    if st.session_state.flag:
        st.download_button(
            label="Last ned prosjektregnskap zip :rocket:",
            data=download_all(
                st.session_state.year,
                st.session_state.selskapsnavn,
                st.session_state.prosjekttittel,
                st.session_state.prosjektnummer,
                st.session_state.prosjektleder,
                st.session_state.antall_ansatte,
                st.session_state.antall_arbeidspakker,
                st.session_state.prosjektbeskrivelse,
                st.session_state.godkjenningsperiode,
                st.session_state.prosjektstatusdato,
                st.session_state.prosjektstatus,
                st.session_state.godkjentdato,
                st.session_state.kategori,
                st.session_state.selskapvansker,
                st.session_state.kostnadsalternativer,
                st.session_state.erkleringdato
            ),
            file_name="SF25 - Prosjektregnskap.zip",  # Added .zip extension
            mime="application/zip",  # Correct MIME type
            key="download6"
        )