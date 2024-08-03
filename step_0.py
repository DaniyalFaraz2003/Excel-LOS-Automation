import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
from datetime import datetime
from copy import copy

def copy_cell_styles(source_cell, target_cell):
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)


def step_0():
    wbOrg = openpyxl.load_workbook("./example_los/Example0_LOS Original.xlsx")
    wbOut = openpyxl.load_workbook("./bot_outputs/step_0.xlsx")
    wsOut = wbOut.active
    wsOrg = wbOrg.active  # active worksheet
    # extract all names from first column to find total number of records
    first_column = [cell.value for cell in wsOrg['A'][1:]]
    names = list({k: None for k in first_column}.keys())

    number_of_records = len(names)

    ALL_RECORDS = {
        "Volumes": [
            "Oil Sales - Bbls",
            "Gas Sales - mcf",
            "NGL Sales - Bbls",
            "NGL Sales - Gal"
        ],
        "Revenue": [
            "Oil Sales Rev",
            "Gas Sales Rev",
            "NGL Sales Rev",
            "Oil Rev Deduct",
            "Gas Rev Deduct",
            "NGL Rev Deduct"
        ],
        "Operating Expenses": [
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
        
    
    for name in names:
        for category_heading in ALL_RECORDS.keys():
            pass
            for category in ALL_RECORDS[category_heading]:
                pass
    wbOut.save("./bot_outputs/step_0_out.xlsx")


step_0()
