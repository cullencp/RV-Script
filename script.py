import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os
from tkinter import PhotoImage

def detect_column_indices(sheet, header_row=1, synonyms=None):
    """Detect column indices dynamically based on the header names."""
    headers = {str(cell.value).strip().lower(): cell.column for cell in sheet[header_row] if cell.value}
    column_map = {}
    for field, variations in synonyms.items():
        for variation in variations:
            if variation.lower() in headers:
                column_map[field] = headers[variation.lower()]
                break
    return column_map

def format_value(value):
    """Format the value in uppercase."""
    if value is None or (isinstance(value, str) and value.strip().lower() == "n/a"):
        return "N/A"
    if isinstance(value, str):
        return value.upper()
    return value

def generate_rv_forms(input_file, output_file, project, client, reference_document, document_revision, start_row, template_type, progress_var, log_file):
    try:
        # Load the workbook
        wb = openpyxl.load_workbook(input_file)
        sheet1 = wb[wb.sheetnames[0]]

        # Define header synonyms for dynamic column detection
        header_synonyms = {
            "Instrument": {
                "Tag": ["Instrument Tag", "Tag"],
                "Make": ["Manufacturer", "Instrument Manufacturer"],
                "Model": ["Model", "Instrument Model"],
                "Order Code": ["Order Code"],
                "Process Connection": ["Process Connection"],
                "Signal Type": ["Control Signal"],
                "Min": ["Range Min", "Min Range"],
                "Max": ["Range Max", "Max Range"],
                "Unit": ["Unit"]
            },
            "Valve": {
                "Instrument Tag": ["BMS Tag", "Tag", "Instrument Tag"],
                "Valve Make Model Number": ["Valve Model Number"],
                "Actuator Make Model Number": ["Actuator Model Number"],
                "Process Connection": ["Process Connection"],
                "Line Size": ["Valve Size[mm]"],
                "Signal Type": ["Actuator Control Signal"],
                "Dial Setting": ["Dial Setting"],
                "Flow Rate": ["Flow Rate"]
            }
        }

        column_map = detect_column_indices(sheet1, header_row=1, synonyms=header_synonyms[template_type])
        template_sheet = wb[wb.sheetnames[-1]]  # Use the last sheet as template
        current_date = datetime.now().strftime("%d %b %Y").upper()

        for index, row in enumerate(sheet1.iter_rows(min_row=start_row, values_only=True), start=1):
            if not row[0]:
                continue

            field_values = {
                field: format_value(row[column_map[field] - 1]) if field in column_map and column_map[field] else None
                for field in column_map
            }

            new_sheet = wb.copy_worksheet(template_sheet)
            rv_form_name = f"RV{str(index).zfill(2)}"
            new_sheet.title = rv_form_name

            if template_type == "Instrument":
                new_sheet["A11"] = field_values["Tag"]
                new_sheet["C11"] = field_values["Make"]
                new_sheet["E11"] = field_values["Model"]
                new_sheet["A14"] = field_values["Process Connection"]
                new_sheet["B14"] = field_values["Order Code"]
                new_sheet["D14"] = field_values["Signal Type"]
                new_sheet["F14"] = field_values["Min"]
                new_sheet["G14"] = field_values["Max"]
                new_sheet["H14"] = field_values["Unit"]
                new_sheet["I14"] = "N/A"
            else:
                new_sheet["A11"] = field_values["Instrument Tag"]
                new_sheet["C11"] = field_values["Valve Make Model Number"]
                new_sheet["F11"] = field_values["Actuator Make Model Number"]
                new_sheet["A15"] = field_values["Process Connection"]
                new_sheet["B15"] = field_values["Line Size"]
                new_sheet["D13"] = field_values["Signal Type"]
                new_sheet["F15"] = field_values["Dial Setting"]
                new_sheet["I15"] = field_values["Flow Rate"]

        wb.save(output_file)
        progress_var.set(100)
        messagebox.showinfo("Success", f"RV forms generated successfully! Output saved as {output_file}")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def main():
    root = tk.Tk()
    root.title("RV Form Generator")
    root.configure(bg="#ffffff")

    input_file_var = tk.StringVar()
    output_folder_var = tk.StringVar()
    project_var = tk.StringVar()
    client_var = tk.StringVar()
    reference_var = tk.StringVar()
    revision_var = tk.StringVar()
    start_row_var = tk.StringVar()
    template_type_var = tk.StringVar(value="Instrument")
    progress_var = tk.DoubleVar()

    logo_image = PhotoImage(file="subnetlogo.png")
    logo_label = tk.Label(root, image=logo_image, bg="#ffffff")
    logo_label.grid(row=0, column=0, columnspan=3, pady=(10, 20))

    tk.Label(root, text="Input File:", bg="#ffffff", font=("Arial", 10)).grid(row=1, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(root, textvariable=input_file_var, width=50).grid(row=1, column=1, padx=5, pady=5)
    tk.Button(root, text="Browse", command=lambda: input_file_var.set(filedialog.askopenfilename())).grid(row=1, column=2, padx=5, pady=5)

    tk.Label(root, text="Output Folder:", bg="#ffffff", font=("Arial", 10)).grid(row=2, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(root, textvariable=output_folder_var, width=50).grid(row=2, column=1, padx=5, pady=5)
    tk.Button(root, text="Browse", command=lambda: output_folder_var.set(filedialog.askdirectory())).grid(row=2, column=2, padx=5, pady=5)

    tk.Button(root, text="Generate RV Forms", command=lambda: generate_rv_forms(input_file_var.get(), os.path.join(output_folder_var.get(), "output.xlsx"), project_var.get(), client_var.get(), reference_var.get(), revision_var.get(), int(start_row_var.get() or 6), template_type_var.get(), progress_var, "log.txt")).grid(row=3, column=0, columnspan=3, pady=15)

    root.mainloop()

if __name__ == "__main__":
    main()
