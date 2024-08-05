import openpyxl
from openpyxl.styles import Font, numbers
from openpyxl.utils import get_column_letter
from step_0 import copy_cell_styles

def add_btu_tab(wbIn):
    if "Example0gross_BTU" not in wbIn.worksheets:
        # only create the sheet if it is not present
        wbIn.create_sheet("Example0gross_BTU")
    ws_btu = wbIn.worksheets[4]  # the sheet that will contain BTU
    # clear the sheet if it is not newly created before copying contents
    for row in ws_btu.iter_rows():
        for cell in row:
            cell.value = None

    # to copy the contents into the main workbook tab, open the source from open refine outputs
    wb_btu_source = openpyxl.load_workbook("./example_los/OpenRefine Outputs/Example0gross_BTU.xlsx")
    ws_btu_source = wb_btu_source.active  # source from which to be copied

    for row in ws_btu_source.iter_rows(values_only=True):
        ws_btu.append(row)

    # correct formatting
    for i in range(1, ws_btu.max_row + 1):
        for j in range(1, ws_btu.max_column + 1):
            copy_cell_styles(ws_btu_source.cell(i, j), ws_btu.cell(i, j))

    # fix top row for btu tab
    ws_btu.freeze_panes = 'A2'

    # bold the headings
    ws_btu['A1'].font = ws_btu['B1'].font = ws_btu['C1'].font = Font(bold=True)

    wb_btu_source.close()  # close the source workbook as it is no longer needed

def add_historical_prod_tab(wbIn):
    if "Example0gross_HistoricalProd" not in wbIn.worksheets:
        # only create the sheet if it is not present
        wbIn.create_sheet("Example0gross_HistoricalProd")
    ws_historical_prod = wbIn.worksheets[5]  # the sheet that will contain Historical Productions
    # clear the sheet if it is not newly created before copying contents
    for row in ws_historical_prod.iter_rows():
        for cell in row:
            cell.value = None

    # to copy the contents into the main workbook tab, open the source from open refine outputs
    wb_historical_prod_source = openpyxl.load_workbook("./example_los/OpenRefine Outputs/Example0gross_HistoricalProduction.xlsx")
    ws_historical_prod_source = wb_historical_prod_source.active  # source from which to be copied

    for row in ws_historical_prod_source.iter_rows(values_only=True):
        ws_historical_prod.append(row)

    # correct formatting
    for i in range(1, ws_historical_prod.max_row + 1):
        for j in range(1, ws_historical_prod.max_column + 1):
            copy_cell_styles(ws_historical_prod_source.cell(i, j), ws_historical_prod.cell(i, j))

    # fix top row for btu tab
    ws_historical_prod.freeze_panes = 'A2'

    # bold the headings
    ws_historical_prod['A1'].font = ws_historical_prod['B1'].font = ws_historical_prod['C1'].font = Font(bold=True)
    ws_historical_prod['D1'].font = ws_historical_prod['E1'].font = ws_historical_prod['F1'].font = ws_historical_prod['G1'].font = Font(bold=True)

    wb_historical_prod_source.close()  # close the source workbook as it is no longer needed

def populate_btu(wsIn):
    i = 52
    while i <= wsIn.max_row:
        wsIn.cell(i, 5).value = f'=IFERROR(VLOOKUP($A{i},Example0gross_BTU!$B:$C,2,0),"")'
        wsIn.cell(i, 5).number_format = '0.000'
        i += 85 # next btu comes after every 85 rows

def populate_first_part(wsIn):
    i = 54
    while i <= wsIn.max_row:
        min = i - 49
        max = i - 6
        for j in range(5, 17): # months
            letter = get_column_letter(j)
            wsIn.cell(i, j).value = f'=IFERROR((SUMIF($D{min}:$D{max},"Oil Sales Revenue ($)",{letter}{min}:{letter}{max})-ABS(SUMIF($D{min}:$D{max},"Oil Revenue Deductions ($)",{letter}{min}:{letter}{max})))/SUMIF($D{min}:$D{max},"Oil Sales Volumes (bbl)",{letter}{min}:{letter}{max}),"")'
            wsIn.cell(i, j).number_format = '0.00'
        i += 85 # next part comes after every 85 rows

    i = 55
    while i <= wsIn.max_row:
        min = i - 50
        max = i - 7
        btu = i - 3
        for j in range(5, 17):  # months
            letter = get_column_letter(j)
            wsIn.cell(i, j).value = f'=IFERROR(((SUMIF($D{min}:$D{max},"Gas Sales Revenue ($)",{letter}{min}:{letter}{max})-ABS(SUMIF($D{min}:$D{max},"Gas Revenue Deductions ($)",{letter}{min}:{letter}{max})))/SUMIF($D{min}:$D{max},"Gas Sales Volumes (mcf)",{letter}{min}:{letter}{max}))/$E{btu},"")'
            wsIn.cell(i, j).number_format = '0.00'
        i += 85  # next part comes after every 85 rows

    i = 56
    while i <= wsIn.max_row:
        min = i - 51
        max = i - 8
        for j in range(5, 17): # months
            letter = get_column_letter(j)
            wsIn.cell(i, j).value = f'=IFERROR((SUMIF($D{min}:$D{max},"NGL Sales Revenue ($)",{letter}{min}:{letter}{max})-ABS(SUMIF($D{min}:$D{max},"NGL Revenue Deductions ($)",{letter}{min}:{letter}{max})))/(SUMIF($D{min}:$D{max},"NGL Sales Volumes (bbl)",{letter}{min}:{letter}{max})+(SUMIF($D{min}:$D{max},"NGL Sales Volumes (gal)",{letter}{min}:{letter}{max})/42)),"")'
            wsIn.cell(i, j).number_format = '0.00'
        i += 85  # next part comes after every 85 rows

def populate_second_part(wsIn):
    i = 58
    while i <= wsIn.max_row:
        value = i - 4
        for j in range(5, 17): # months
            letter = get_column_letter(j)
            wsIn.cell(i, j).value = f'=IFERROR({letter}{value}-{letter}$1,"")'
            wsIn.cell(i, j).number_format = numbers.BUILTIN_FORMATS[40] # to display negative numbers in red and brackets
        num = 14
        for j in range(17, 21): # averages
            letter = get_column_letter(num)
            wsIn.cell(i, j).value = f'=IFERROR(AVERAGE({letter}{i}:P{i}),"")'
            wsIn.cell(i, j).number_format = numbers.BUILTIN_FORMATS[40]
            num -= 3
        i += 85 # next part comes after every 85 rows

    i = 59
    while i <= wsIn.max_row:
        value = i - 4
        for j in range(5, 17): # months
            letter = get_column_letter(j)
            wsIn.cell(i, j).value = f'=IFERROR({letter}{value}-{letter}$2,"")'
            wsIn.cell(i, j).number_format = numbers.BUILTIN_FORMATS[40]
        num = 14
        for j in range(17, 21):  # averages
            letter = get_column_letter(num)
            wsIn.cell(i, j).value = f'=IFERROR(AVERAGE({letter}{i}:P{i}),"")'
            wsIn.cell(i, j).number_format = numbers.BUILTIN_FORMATS[40]
            num -= 3
        i += 85

    i = 60
    while i <= wsIn.max_row:
        value = i - 4
        for j in range(5, 17):  # months
            letter = get_column_letter(j)
            wsIn.cell(i, j).value = f'=IFERROR({letter}{value}-{letter}$1,"")'
            wsIn.cell(i, j).number_format = numbers.BUILTIN_FORMATS[40]
        num = 14
        for j in range(17, 21):  # averages
            letter = get_column_letter(num)
            wsIn.cell(i, j).value = f'=IFERROR(AVERAGE({letter}{i}:P{i}),"")'
            wsIn.cell(i, j).number_format = numbers.BUILTIN_FORMATS[40]
            num -= 3
        i += 85

def populate_third_part(wsIn):
    i = 62
    while i <= wsIn.max_row:
        value = i - 8
        for j in range(5, 17): # months
            letter = get_column_letter(j)
            wsIn.cell(i, j).value = f'=IFERROR({letter}{value}/{letter}$1,"")'
            wsIn.cell(i, j).number_format = '0.00%'
        num = 14
        for j in range(17, 21):  # averages
            letter = get_column_letter(num)
            wsIn.cell(i, j).value = f'=IFERROR(AVERAGE({letter}{i}:P{i}),"")'
            wsIn.cell(i, j).number_format = '0.00%'
            num -= 3
        i += 85

    i = 63
    while i <= wsIn.max_row:
        value = i - 8
        for j in range(5, 17):
            letter = get_column_letter(j)
            wsIn.cell(i, j).value = f'=IFERROR({letter}{value}/{letter}$2,"")'
            wsIn.cell(i, j).number_format = '0.00%'
        num = 14
        for j in range(17, 21):  # averages
            letter = get_column_letter(num)
            wsIn.cell(i, j).value = f'=IFERROR(AVERAGE({letter}{i}:P{i}),"")'
            wsIn.cell(i, j).number_format = '0.00%'
            num -= 3
        i += 85

    i = 64
    while i <= wsIn.max_row:
        value = i - 8
        for j in range(5, 17):
            letter = get_column_letter(j)
            wsIn.cell(i, j).value = f'=IFERROR({letter}{value}/{letter}$1,"")'
            wsIn.cell(i, j).number_format = '0.00%'
        num = 14
        for j in range(17, 21):  # averages
            letter = get_column_letter(num)
            wsIn.cell(i, j).value = f'=IFERROR(AVERAGE({letter}{i}:P{i}),"")'
            wsIn.cell(i, j).number_format = '0.00%'
            num -= 3
        i += 85

def populate_fourth_part(wsIn):
    i = 66
    while i <= wsIn.max_row:
        for j in range(5, 17):
            letter = get_column_letter(j)
            wsIn.cell(i, j).value = f'=XLOOKUP($A{i}&"|"&{letter}$4,Example0gross_HistoricalProd!$B:$B&"|"&Example0gross_HistoricalProd!$C:$C,Example0gross_HistoricalProd!$D:$D,"",0,1)'
            wsIn.cell(i, j).number_format = '0.00'
        i += 85

    # i = 67
    # while i <= wsIn.max_row:
    #     min = i - 62
    #     max = i - 19
    #     value = i - 1
    #     for j in range(5, 17):
    #         letter = get_column_letter(j)
    #         wsIn.cell(i, j).value = f'=IFERROR(SUMIF($D{min}:$D{max},"Gas Sales Volumes (mcf)",{letter}{min}:{letter}{max})/{letter}{value},"")'
    #         wsIn.cell(i, j).number_format = '0.00%'
    #     num = 14
    #     for j in range(17, 21):  # averages
    #         letter = get_column_letter(num)
    #         wsIn.cell(i, j).value = f'=IFERROR(AVERAGE({letter}{i}:P{i}),"")'
    #         wsIn.cell(i, j).number_format = '0.00%'
    #         num -= 3
    #     i += 85

def populate_fifth_part(wsIn):
    pass

def populate_sixth_part(wsIn):
    pass

def populate_seventh_part(wsIn):
    pass

def populate_eighth_part(wsIn):
    pass

def populate_ninth_part(wsIn):
    pass

def populate_last_part(wsIn):
    pass

def step_5():
    wbIn = openpyxl.load_workbook("./bot_outputs/step_4_out.xlsx")
    add_btu_tab(wbIn)
    add_historical_prod_tab(wbIn)

    wsIn = wbIn.worksheets[0] # active LOS worksheet

    # populate values
    populate_btu(wsIn)
    populate_first_part(wsIn)
    populate_second_part(wsIn)
    populate_third_part(wsIn)
    populate_fourth_part(wsIn)
    populate_fifth_part(wsIn)
    populate_sixth_part(wsIn)
    populate_seventh_part(wsIn)
    populate_eighth_part(wsIn)
    populate_ninth_part(wsIn)
    populate_last_part(wsIn)

    wbIn.save("./bot_outputs/step_5_out.xlsx")

step_5()