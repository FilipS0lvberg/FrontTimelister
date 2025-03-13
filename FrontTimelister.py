import streamlit as st
from supporting_functions import *
from create_workbook import create_workbook1


st.title("Front Generation - Timelister :memo:")


year = st.number_input("Årstall", min_value=2025, max_value=2035, value=2025)
selskapsnavn = st.text_input("Selskapsnavn")
prosjekttittel = st.text_input("Prosjekttittel")
prosjektleder = st.text_input("Prosjektleder")
prosjektnummer = st.number_input("Prosjektnummer")
ant_ansatte = st.number_input("Antall ansatte", value=1, min_value=1)
antall_arbeidspakker = st.number_input("Antall Arbeidspakker", value=1, min_value=1)
prosjektbeskrivelse = st.text_area("Prosjektbeskrivelse", height=700)

#ansatte = get_workers(ant_ansatte)

#kvartaler = {}
#wbio = create_workbook1(year=year, quarter=kvartal, selskap=selskapsnavn, prosjekttittel=prosjekttittel, prosjektnummer=prosjektnummer, prosjektleder=prosjektleder, prosject_description=prosjektbeskrivelse, workers=ansatte, num_workpackages=antall_arbeidspakker)

####################### DOWNLOAD BUTTON FOR DOWNLOADING WORD DOC
st.download_button(
    type="primary",
    label="Last ned Q1 - Timelister",
    # data=write_docx(st.session_state.norsk_proj_title, st.session_state.main_checkboxes ,st.session_state.main_box_contents, st.session_state.workpackages),
    data=create_workbook1(year=year, quarter=1, selskap=selskapsnavn, prosjekttittel=prosjekttittel, prosjektnummer=prosjektnummer, prosjektleder=prosjektleder, prosject_description=prosjektbeskrivelse, workers=get_workers(ant_ansatte), num_workpackages=antall_arbeidspakker),
    file_name=f"SF{year} - Timelister Q1 - {selskapsnavn}",
    mime="xlsx"
)

st.download_button(
    type="primary",
    label="Last ned Q2 - Timelister",
    # data=write_docx(st.session_state.norsk_proj_title, st.session_state.main_checkboxes ,st.session_state.main_box_contents, st.session_state.workpackages),
    data=create_workbook1(year=year, quarter=2, selskap=selskapsnavn, prosjekttittel=prosjekttittel, prosjektnummer=prosjektnummer, prosjektleder=prosjektleder, prosject_description=prosjektbeskrivelse, workers=get_workers(ant_ansatte), num_workpackages=antall_arbeidspakker),
    file_name=f"SF{year} - Timelister Q2 - {selskapsnavn}",
    mime="xlsx"
)

st.download_button(
    type="primary",
    label="Last ned Q3 - Timelister",
    # data=write_docx(st.session_state.norsk_proj_title, st.session_state.main_checkboxes ,st.session_state.main_box_contents, st.session_state.workpackages),
    data=create_workbook1(year=year, quarter=3, selskap=selskapsnavn, prosjekttittel=prosjekttittel, prosjektnummer=prosjektnummer, prosjektleder=prosjektleder, prosject_description=prosjektbeskrivelse, workers=get_workers(ant_ansatte), num_workpackages=antall_arbeidspakker),
    file_name=f"SF{year} - Timelister Q3 - {selskapsnavn}",
    mime="xlsx"
)

st.download_button(
    type="primary",
    label="Last ned Q4 - Timelister",
    # data=write_docx(st.session_state.norsk_proj_title, st.session_state.main_checkboxes ,st.session_state.main_box_contents, st.session_state.workpackages),
    data=create_workbook1(year=year, quarter=4, selskap=selskapsnavn, prosjekttittel=prosjekttittel, prosjektnummer=prosjektnummer, prosjektleder=prosjektleder, prosject_description=prosjektbeskrivelse, workers=get_workers(ant_ansatte), num_workpackages=antall_arbeidspakker),
    file_name=f"SF{year} - Timelister Q4 - {selskapsnavn}",
    mime="xlsx"
)
