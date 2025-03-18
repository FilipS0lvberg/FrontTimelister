from openpyxl.workbook import Workbook
from openpyxl.styles import Font, Color, Alignment, Border, Side, PatternFill
from supporting_functions import get_dates_in_quarter, get_first_letters_dict
import io

def create_prosjektregnskap1(year: int, juridisknavn: str, project_number: int, project_title: str, godkjent_periode: str, status_dato: str, status: str, godkjent_dato: str, kategori: str, selskap_vansker: str, cost_options):
    wb = Workbook()

    ###################### CREATE FRONT PAGE (REPORT) (Sheet 0) ##########################
    ws0 = wb.active
    ws0.sheet_view.showGridLines = False
    ws0.title = "Rapport"

    # Light gray
    fill_lg = PatternFill(start_color='00C0C0C0', 
            end_color='00C0C0C0', 
            fill_type='solid')

    ws0["B2"] = f"Prosjektregnskap {year}"
    ws0["B2"].font = Font(size=20, bold=True)

    ws0.column_dimensions["B"].width = 50
    ws0.column_dimensions["C"].width = 50


    ws0["B4"] = "Juridisk selskapsnavn"
    ws0["B5"] = "Prosjektnummer"
    ws0["B6"] = "Prosjekttittel (Beskrivelse)"

    ws0["B8"] = "Godkjenningsperiode (dato fra - dato til)"
    ws0["B9"] = "Prosjektstatusdato"
    ws0["B10"] = "Prosjektstatus"
    ws0["B11"] = "Søknad godkjent dato"
    ws0["B12"] = "Kategori"
    ws0["B13"] = "Var selskapet i vansker ved godkjenningstidspunkt målt etter siste avlagte årsregnskap?"
    ws0["B13"].alignment = Alignment(wrap_text=True)
    
    ### C kolonne
    
    ws0["C4"] = juridisknavn
    ws0["C4"].alignment = Alignment(horizontal="right")
    ws0["C5"] = project_number
    ws0["C5"].alignment = Alignment(horizontal="right")
    ws0["C6"] = project_title
    ws0["C6"].alignment = Alignment(horizontal="right")
    ws0["C8"] = godkjent_periode
    ws0["C8"].alignment = Alignment(horizontal="right")
    ws0["C9"] = status_dato
    ws0["C9"].alignment = Alignment(horizontal="right")
    ws0["C10"] = status
    ws0["C10"].alignment = Alignment(horizontal="right")
    ws0["C11"] = godkjent_dato
    ws0["C11"].alignment = Alignment(horizontal="right")
    ws0["C12"] = kategori
    ws0["C12"].alignment = Alignment(horizontal="right")
    ws0["C13"] = selskap_vansker
    ws0["C13"].alignment = Alignment(horizontal="right")

    # Adding borders
    start_cell = "B4"
    end_cell = "C6"
    cell_range = ws0[start_cell:end_cell]

    border_A = Border(bottom=Side(style="thin"), top=Side(style="thin"), left=Side(style="thin"), right=Side(style="thin"))
    for row in cell_range:
        for cell in row:
            cell.border = border_A

    start_cell = "B8"
    end_cell = "C13"
    cell_range = ws0[start_cell:end_cell]

    for row in cell_range:
        for cell in row:
            cell.border = border_A


    # Begynnelse av Rapport sammendrag
    ws0["B16"] = "Prosjektkostnader"
    ws0["B16"].font = Font(size=12, bold=True)

    border_u = Border(bottom=Side(style="thin"))
    ws0["B16"].border = border_u
    ws0["c16"].border = border_u

    # TODO DETTE MÅ GJØRES NOE MED FOR HVER OPTION
    ws0["B17"] = "Timer"
    ws0["B18"] = "Prosjektkostnader"
    ws0["B19"] = "Bruk av eget utstyr"
    ws0["B20"] = "Kapitalkostnader"
    ws0["B20"].border = border_u
    ws0["C20"].border = border_u

    ws0["B21"] = "Sum prosjektkostnader"
    ws0["B21"].font = Font(size=12, bold=True)
    ws0["C21"] = "=SUM(C17:C20)"
    ws0["C21"].font = Font(size=12, bold=True)

    ws0["B23"].border = border_u
    ws0["C23"].border = border_u
    ws0["B24"] = "Grunnlag for støtte - overføres til skattebergningen"
    ws0["B24"].font = Font(size=12, bold=True)
    ws0["C24"] = "=MIN(C21,25000000)"
    ws0["C24"].font = Font(size=12, bold=True)
    ws0["B25"] = "Støttesats"
    ws0["C25"] = "19 %"
    ws0["C25"].alignment = Alignment(horizontal="right")

    ws0["B26"].border = border_u
    ws0["C26"].border = border_u
    ws0["B27"] = "Skattefradrag"
    ws0["B27"].font = Font(size=12, bold=True)
    ws0["C27"] = "=C24*0.19"
    ws0["C27"].font = Font(size=12, bold=True)
    ws0["B27"].border = border_u
    ws0["C27"].border = border_u



    ###########################################################################################################
    #### Time kostnader ark
    if "Timekostnader" in cost_options:
        ws1 = wb.create_sheet("Timekostnader")
        ws1.sheet_view.showGridLines = False

        ws1["B2"] = f"Timekostnader"
        ws1["B2"].font = Font(size=20, bold=True)

        ws1.column_dimensions['B'].width = 30
        ws1.column_dimensions['C'].width = 20
        ws1.column_dimensions['D'].width = 15
        ws1.column_dimensions['E'].width = 15
        ws1.column_dimensions['F'].width = 15

        ws1["B4"] = "Selskap"
        ws1["F4"] = juridisknavn
        ws1["F4"].align = Alignment(horizontal="right")

        ws1["B5"] = "Prosjektnummer"
        ws1["F5"] = project_number
        ws1["F5"].align = Alignment(horizontal="right")

        ws1["B6"] = "Prosjekttittel"
        ws1["F6"] = project_title
        ws1["F6"].align = Alignment(horizontal="right")

        start_cell = "B4"
        end_cell = "F6"
        cell_range = ws1[start_cell:end_cell]

        for row in cell_range:
            for cell in row:
                cell.border = border_u


        # Timer table
        ws1["B9"] = "Navn"
        ws1["B9"].font = Font(size=12, bold=True)
        ws1["B9"].fill = fill_lg

        ws1["C9"] = "Brutto årslønn"
        ws1["C9"].font = Font(size=12, bold=True)
        ws1["C9"].fill = fill_lg

        ws1["D9"] = "Timepris"
        ws1["D9"].font = Font(size=12, bold=True)
        ws1["D9"].fill = fill_lg

        ws1["E9"] = "Timer"
        ws1["E9"].font = Font(size=12, bold=True)
        ws1["E9"].fill = fill_lg

        ws1["F9"] = "Timekostnad"
        ws1["F9"].font = Font(size=12, bold=True)
        ws1["F9"].fill = fill_lg

        for i in range(15):
            ws1[f"D{10+i}"] = f"=MIN(C{10+i}*0.0012,700)"
            ws1[f"D{10+i}"].alignment = Alignment(horizontal="right")
            ws1[f"F{10+i}"] = f"=E{10+i}*D{10+i}"
            ws1[f"F{10+i}"].alignment = Alignment(horizontal="right")
            ws1[f"F{10+i}"].number_format = '#,##0.00kr' 


        ws1["B25"] = "Sum"
        ws1["B25"].fill = fill_lg
        ws1["C25"].fill = fill_lg
        ws1["D25"].fill = fill_lg
        ws1["E25"] = "=SUM(E10:E24)"
        ws1["E25"].fill = fill_lg
        ws1["F25"] = "=SUM(F10:F24)"
        ws1["F25"].number_format = '#,##0.00kr'
        ws1["F25"].fill = fill_lg

        start_cell = "B9"
        end_cell = "F25"
        cell_range = ws1[start_cell:end_cell]

        for row in cell_range:
            for cell in row:
                cell.border = border_A

        ws1["B28"] = "*Brutto årslønn er kun spesifisert dersom den er under 584 000 kr/år"
        ws1["B29"] = "** Anvendt årslønn er oppgitt etter siste lønnskorrigering før 31/12 inneværende skatteår, ref FSFIN § 16-40-6 tredje ledd."



    # page for prosjektkostnader
    #if "Prosjektkostnader" in cost_options:




    # save workbook
    wb.save("testbook.xlsx")
    # bio = io.BytesIO()
    # wb.save(bio)
    # return bio.getvalue()



if __name__ == "__main__":
    cost_opts = ["Timekostnader", "Prosjektkostnader", "Bruk av eget utstyr", "Kapitalkostnader"]
    create_prosjektregnskap1(2025, "Front Innovation AS", 1283, "Nytt test prosjekt", godkjent_periode="01.01.2025 - 01.01.2026", status_dato="01.02.2025", status="Godkjent", godkjent_dato="06.09.2024", kategori="Utviklingsprosjekt eksperimentell utvikling", selskap_vansker="Nei", cost_options=cost_opts)