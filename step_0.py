import openpyxl

def step_0():
    wb = openpyxl.load_workbook("./Example0_LOS Original.xlsx")
    ws = wb.active # active worksheet
    # extract all names from first column to find total number of records
    first_column = [cell.value for cell in ws['A'][1:]]
    names = list({k: None for k in first_column}.keys())

    number_of_records = len(names)
    


    wb.save("./bot_outputs/step_0.xlsx")