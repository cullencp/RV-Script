import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from datetime import datetime

def generate_rv_forms(input_file, output_file):
    # Load the workbook
    wb = openpyxl.load_workbook(input_file)

    # Get the sheets
    sheet1 = wb[wb.sheetnames[0]]  # Parent data sheet (Sheet 1)
    template_sheet = wb["RV Instrument  SUB-TF-01"]  # Template (Sheet 3)

    # User inputs
    project = input("Enter Project Name: ")
    client = input("Enter Client Name: ")
    reference_document = input("Enter Reference Document: ")
    document_revision = input("Enter Document Revision: ")
    start_row = int(input("Enter the starting row for instruments (e.g., 6, 7, 8): "))

    # Current date in the required format
    current_date = datetime.now().strftime("%d %b %Y")

    # Loop through rows in Sheet 1 starting at the user-defined row
    for index, row in enumerate(sheet1.iter_rows(min_row=start_row, values_only=True), start=1):
        # Get the data from the row
        instrument_tag = row[1]  # Column B
        manufacturer = row[4]   # Column E
        model = row[5]          # Column F
        process_connection = row[6]  # Column G
        immersion_length = row[10]  # Column K
        control_signal = row[18]    # Column S
        min_range = row[13]         # Column N
        max_range = row[14]         # Column O
        unit = row[15]              # Column P

        # Create a new sheet for each instrument
        rv_form_name = f"RV{str(index).zfill(2)}"
        new_sheet = wb.copy_worksheet(template_sheet)
        new_sheet.title = rv_form_name

        # Populate static fields
        new_sheet["A5"] = project
        new_sheet["E5"] = client
        new_sheet["A7"] = reference_document
        new_sheet["E7"] = document_revision
        new_sheet["I5"] = current_date
        new_sheet["I7"] = rv_form_name

        # Populate dynamic fields
        new_sheet["A11"] = instrument_tag
        new_sheet["C11"] = manufacturer
        new_sheet["E11"] = model
        new_sheet["A14"] = process_connection
        new_sheet["B14"] = immersion_length
        new_sheet["D14"] = control_signal
        new_sheet["F14"] = min_range
        new_sheet["G14"] = max_range
        new_sheet["H14"] = unit
        new_sheet["I14"] = "N/A"

        # Apply alignment to all populated cells
        for cell_ref in ["A5", "E5", "A7", "E7", "I5", "I7", "A11", "C11", "E11", "A14", "B14", "D14", "F14", "G14", "H14", "I14"]:
            cell = new_sheet[cell_ref]
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # Save the updated workbook
    wb.save(output_file)

# Usage
input_file = "input165.xlsx"  # Replace with your input file
output_file = "output165.xlsx"  # Replace with your desired output file
generate_rv_forms(input_file, output_file)

