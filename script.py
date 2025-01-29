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

        # Get the data sheet (Sheet 1)
        sheet1 = wb[wb.sheetnames[0]]

        # Define header synonyms for instrument and valve templates
        header_synonyms = {
            "Instrument": {
                "Instrument Tag": ["Instrument Tag", "Tag"],
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

        # Detect column indices dynamically based on the selected template type
        column_map = detect_column_indices(sheet1, header_row=1, synonyms=header_synonyms[template_type])

        # Get the pre-loaded template sheet
        template_sheet_name = get_sheet_by_partial_name(wb, f"RV {template_type}")
        template_sheet = wb[template_sheet_name]

        # Current date in the required format
        current_date = datetime.now().strftime("%d %b %Y").upper()

        # Open log file
        with open(log_file, "a") as log:
            log.write(f"\n{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Starting RV form generation for project: {project}\n")

            # Get the total number of rows to process
            total_rows = sum(1 for _ in sheet1.iter_rows(min_row=start_row, values_only=True) if _[0])
            progress_step = 100 / total_rows if total_rows > 0 else 100
            progress = 0

            # Loop through rows in Sheet 1 starting at the user-defined row
            for index, row in enumerate(sheet1.iter_rows(min_row=start_row, values_only=True), start=1):
                if not row[0]:
                    continue

                # Dynamically retrieve data based on column_map
                field_values = {
                    field: format_value(row[column_map[field] - 1]) if field in column_map and column_map[field] else None
                    for field in column_map
                }

                # Create a new sheet for each instrument/valve
                new_sheet = wb.copy_worksheet(template_sheet)
                rv_form_name = f"RV{str(index).zfill(2)}"
                new_sheet.title = rv_form_name

                # Populate dynamic fields based on template type
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
                else:  # Valve Template
                    new_sheet["A11"] = field_values["Instrument Tag"]
                    new_sheet["C11"] = field_values["Valve Make Model Number"]
                    new_sheet["F11"] = field_values["Actuator Make Model Number"]
                    new_sheet["A15"] = field_values["Process Connection"]
                    new_sheet["B15"] = field_values["Line Size"]
                    new_sheet["D13"] = field_values["Signal Type"]
                    new_sheet["F15"] = field_values["Dial Setting"]
                    new_sheet["I15"] = field_values["Flow Rate"]

                # Populate static fields for both templates
                new_sheet["A5"] = project
                new_sheet["E5"] = client
                new_sheet["A7"] = reference_document
                new_sheet["E7"] = document_revision
                new_sheet["I5"] = current_date
                new_sheet["I7"] = rv_form_name

                # Apply alignment and font size to all populated cells
                font = Font(size=11)  # Size 11 for all text
                for cell_ref in [
                    "A5", "E5", "A7", "E7", "I5", "I7", "A11", "C11", "E11", "A14", "B14", "D14", "F14", "G14", "H14", "I14",
                    "F11", "A15", "B15", "D13", "F15", "I15"
                ]:
                    cell = new_sheet[cell_ref]
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    cell.font = font

                # Log the processed RV form
                log.write(f"Processed: {rv_form_name}\n")

                # Update progress bar
                progress += progress_step
                progress_var.set(progress)

            # Save the updated workbook
            wb.save(output_file)
            progress_var.set(100)  # Ensure progress reaches 100%
            log.write("RV form generation completed successfully.\n")
            messagebox.showinfo("Success", f"RV forms generated successfully! Output saved as {output_file}")

    except Exception as e:
        with open(log_file, "a") as log:
            log.write(f"Error: {str(e)}\n")
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def get_sheet_by_partial_name(wb, partial_name):
    for sheet in wb.sheetnames:
        if partial_name.lower() in sheet.lower():
            return sheet
    raise ValueError(f"Sheet with partial name '{partial_name}' not found.")

# GUI Implementation
def main():
    def browse_input_file():
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xlsm"), ("All files", "*.*")])
        input_file_var.set(file_path)

    def browse_output_location():
        folder_path = filedialog.askdirectory()
        output_folder_var.set(folder_path)

    def generate_forms():
        input_file = input_file_var.get()
        project = project_var.get()
        client = client_var.get()
        reference_document = reference_var.get()
        document_revision = revision_var.get()
        start_row = int(start_row_var.get() or 6)  # Default to row 6
        output_folder = output_folder_var.get()
        template_type = template_type_var.get()

        if not input_file or not project or not client or not reference_document or not document_revision or not output_folder or not template_type:
            messagebox.showerror("Input Error", "Please fill in all fields and select files.")
            return

        # Generate default output file name based on project name
        file_extension = os.path.splitext(input_file)[-1]
        output_file = os.path.join(output_folder, f"{project}-RVs{file_extension}")

        log_file = os.path.join(output_folder, "rv_generator_log.txt")

        try:
            generate_rv_forms(input_file, output_file, project, client, reference_document, document_revision, start_row, template_type, progress_var, log_file)
        except Exception as e:
            with open(log_file, "a") as log:
                log.write(f"Error: {str(e)}\n")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    root = tk.Tk()
    root.title("RV Form Generator")
    root.configure(bg="#ffffff")  # Set background to white for a clean look

    # Variables
    input_file_var = tk.StringVar()
    output_folder_var = tk.StringVar()
    project_var = tk.StringVar()
    client_var = tk.StringVar()
    reference_var = tk.StringVar()
    revision_var = tk.StringVar()
    start_row_var = tk.StringVar()
    template_type_var = tk.StringVar(value="Instrument")
    progress_var = tk.DoubleVar()

    # Add the logo to the GUI
    logo_image = PhotoImage(file="subnetlogo.png")
    logo_label = tk.Label(root, image=logo_image, bg="#ffffff")
    logo_label.grid(row=0, column=0, columnspan=3, pady=(10, 20))  # Adjust padding for spacing

    # GUI Layout
    tk.Label(root, text="Input File:", bg="#ffffff", fg="#333333", font=("Arial", 10)).grid(row=1, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(root, textvariable=input_file_var, width=50).grid(row=1, column=1, padx=5, pady=5)
    tk.Button(root, text="Browse", command=browse_input_file, bg="#4CAF50", fg="#ffffff", font=("Arial", 10)).grid(row=1, column=2, padx=5, pady=5)

    tk.Label(root, text="Output Location:", bg="#ffffff", fg="#333333", font=("Arial", 10)).grid(row=2, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(root, textvariable=output_folder_var, width=50).grid(row=2, column=1, padx=5, pady=5)
    tk.Button(root, text="Browse", command=browse_output_location, bg="#4CAF50", fg="#ffffff", font=("Arial", 10)).grid(row=2, column=2, padx=5, pady=5)

    tk.Label(root, text="Template Type:", bg="#ffffff", fg="#333333", font=("Arial", 10)).grid(row=3, column=0, padx=5, pady=5, sticky="e")
    template_dropdown = ttk.Combobox(root, textvariable=template_type_var, values=["Instrument", "Valve"], state="readonly")
    template_dropdown.grid(row=3, column=1, padx=5, pady=5)

    tk.Label(root, text="Project Name:", bg="#ffffff", fg="#333333", font=("Arial", 10)).grid(row=4, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(root, textvariable=project_var, width=50).grid(row=4, column=1, padx=5, pady=5)

    tk.Label(root, text="Client Name:", bg="#ffffff", fg="#333333", font=("Arial", 10)).grid(row=5, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(root, textvariable=client_var, width=50).grid(row=5, column=1, padx=5, pady=5)

    tk.Label(root, text="Reference Document:", bg="#ffffff", fg="#333333", font=("Arial", 10)).grid(row=6, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(root, textvariable=reference_var, width=50).grid(row=6, column=1, padx=5, pady=5)

    tk.Label(root, text="Document Revision:", bg="#ffffff", fg="#333333", font=("Arial", 10)).grid(row=7, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(root, textvariable=revision_var, width=50).grid(row=7, column=1, padx=5, pady=5)

    tk.Label(root, text="Starting Row (Auto-Detect if Blank):", bg="#ffffff", fg="#333333", font=("Arial", 10)).grid(row=8, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(root, textvariable=start_row_var, width=50).grid(row=8, column=1, padx=5, pady=5)

    tk.Button(root, text="Generate RV Forms", command=generate_forms, bg="#4CAF50", fg="#ffffff", font=("Arial", 12, "bold")).grid(row=9, column=0, columnspan=3, pady=15)

    # Progress Bar
    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
    progress_bar.grid(row=10, column=0, columnspan=3, padx=10, pady=10, sticky="ew")

    # Footer
    tk.Label(root, text="© 2025 Subnet Ltd. All rights reserved.", bg="#ffffff", fg="#333333", font=("Arial", 9)).grid(row=11, column=0, columnspan=3, pady=(10, 0))

    root.mainloop()

if __name__ == "__main__":
    main()
