import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import xlwings as xw

def adjust_column_widths(wsIn):
    app = xw.App(visible=False)
    wb_xlwings = app.books.open('./bot_outputs/step_4_out.xlsx')

    # wb_openpyxl = openpyxl.load_workbook('./bot_outputs/step_4_out.xlsx')
    # ws_openpyxl = wb_openpyxl.active

    sheet_xlwings = wb_xlwings.sheets[0]

    # Adjust column widths based on displayed values
    for col in wsIn.iter_cols(min_row=1, max_row=wsIn.max_row, min_col=1, max_col=wsIn.max_column):
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            value = sheet_xlwings.range((cell.row, cell.column)).value
            if value is not None and len(str(value)) > max_length:
                max_length = len(str(value))
        # adjusted_width = (max_length + 2) * 1.2  # Adjust the multiplier as needed
        wsIn.column_dimensions[col_letter].width = max_length

    # Save and close the workbooks
    # wb_openpyxl.save('./bot_outputs/step_4_out.xlsx')
    wb_xlwings.close()
    app.quit()

def step_4():
    wbIn = openpyxl.load_workbook("./bot_outputs/step_3_out.xlsx")
    
    wsIn = wbIn.worksheets[0]

    columns = list(wsIn.iter_cols(min_col=17, max_col=20, min_row=4, max_row=4))
    for column, month in zip(columns, [3, 6, 9, 12]):
        column[0].value = f"{month}-Mo Avg"
        column[0].font = Font(bold=True)

    # 41 is the number of rows which will now be inserted below every record


    # adjust_column_widths(wsIn)
    wbIn.save("./bot_outputs/step_4_out.xlsx")

step_4()