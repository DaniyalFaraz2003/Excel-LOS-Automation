import openpyxl
from copy import copy

def copy_cell_styles(source_cell, target_cell):
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)


def get_rows_of_name(sheet, search_value):
    matching_rows = []
    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
        if row[0].value == search_value:
            matching_rows.append([cell.value for cell in row])
    
    return matching_rows

def step_0():
    wbOrg = openpyxl.load_workbook("./example_los/Example0_LOS Original.xlsx")
    wbOut = openpyxl.load_workbook("./bot_outputs/step_0_in.xlsx")
    wsOut = wbOut.active
    wsOrg = wbOrg.active  # active worksheet
    # extract all names from first column to find total number of records
    first_column = [cell.value for cell in wsOrg['A'][1:]]
    names = list({k: None for k in first_column}.keys())

    ALL_RECORDS = {
        "Volumes:": [
            "Oil Sales - Bbls",
            "Gas Sales - mcf",
            "NGL Sales - Bbls",
            "NGL Sales - Gal"
        ],
        "Revenue:": [
            "Oil Sales Rev",
            "Gas Sales Rev",
            "NGL Sales Rev",
            "Oil Rev Deduct",
            "Gas Rev Deduct",
            "NGL Rev Deduct"
        ],
        "Operating Expenses:": [
            "Severance Taxes",
            "Other Deductions",
            "Chemicals",
            "Communications",
            "Consulting",
            "Contract Labor",
            "Fuel & Power",
            "Hot Oil & Other Treatments",
            "Insurance",
            "Legal",
            "Marketing",
            "Measurement/Metering",
            "Miscellaneous",
            "Overhead",
            "Professional Services",
            "Pumping & Gauging",
            "Rental Equipment",
            "Repairs & Maintenance",
            "Road & Lease Maintenance",
            "Salt Water Disposal",
            "Supervision",
            "Supplies",
            "Ad Valorem",
            "Trucking & Hauling",
            "Vacuum Truck/Clean Up",
            "Well Servicing",
            "Workover Rig",
            "Gathering & Transport Chg",
            "Swd Disposal Chg",
            "Total Expenses",
            "Net Operating Profit"
        ]
    }
    for c in range(1, wsOrg.max_column + 1):
        wsOut.cell(1, c).value = wsOrg.cell(1, c).value
        copy_cell_styles(wsOrg.cell(1, c), wsOut.cell(1, c))
        
    row = 2
    for name in names:
        name_rows = get_rows_of_name(wsOrg, name)
        for category_heading in ALL_RECORDS.keys():
            wsOut.cell(row, 1).value, wsOut.cell(row, 2).value = name, category_heading
            row += 1
            for category in ALL_RECORDS[category_heading]:
                wsOut.cell(row, 1).value, wsOut.cell(row, 2).value = name, category
                
                found = False
                for _row in name_rows:
                    if _row[1] == category:
                        found = True
                        for c in range(3, wsOrg.max_column + 1):
                            wsOut.cell(row, c).value = _row[c - 1]
                            wsOut.cell(row, c).number_format = '#,##0.00_);[Red](#,##0.00)'
                        break

                if not found:
                    for c in range(3, wsOrg.max_column + 1):
                        wsOut.cell(row, c).value = 0.00
                        wsOut.cell(row, c).number_format = '#,##0.00_);[Red](#,##0.00)'
                row += 1

    # adding filters
    filter_file = openpyxl.load_workbook("./bot_outputs/step_0_filters.xlsx")
    filter_sheet = filter_file.active
    filter_range = filter_sheet.auto_filter.ref
    wsOut.auto_filter.ref = filter_range

    # fix the top row
    wsOut.freeze_panes = 'A2'

    # rename sheet name
    wsOut.title = 'Example0gross_LOS'

    wbOut.save("./bot_outputs/step_0_out.xlsx")


step_0()
