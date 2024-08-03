import openpyxl

def step_0():
    wb = openpyxl.load_workbook("./bot_outputs/step_0.xlsx")
    ws = wb.active # active worksheet
    # extract all names from first column to find total number of records
    first_column = [cell.value for cell in ws['A'][1:]]
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

    


    wb.save("./bot_outputs/step_0.xlsx")

step_0()