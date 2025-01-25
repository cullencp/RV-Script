import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os
from tkinter import PhotoImage

def get_sheet_by_partial_name(wb, partial_name):
    for sheet in wb.sheetnames:
        if partial_name.lower() in sheet.lower():
            return sheet
    raise ValueError(f"Sheet with partial name '{partial_name}' not found.")

def format_value(value):
    """Format the value with first letter capitalized unless it's a full code."""
    if value is None or (isinstance(value, str) and value.strip().lower() == "n/a"):
        return "N/A"
    if isinstance(value, str) and not any(char.islower() for char in value):  # Check if it's a full code
        return value
    return value.capitalize() if isinstance(value, str) else value

def generate_rv_forms(input_file, output_file, project, client, reference_document, document_revision, start_row, template_type, progress_var, log_file):
    try:
        # Load the workbook
        wb = openpyxl.load_workbook(input_file)

        # Get the pre-loaded template sheet dynamically
        if template_type == "Instrument":
            template_sheet_name = get_sheet_by_partial_name(wb, "RV Instrument  SUB-TF-01")
        else:
            template_sheet_name = get_sheet_by_partial_name(wb, "RV Valve SUB-TF-02")

        template_sheet = wb[template_sheet_name]

        # Current date in the required format
        current_date = datetime.now().strftime("%d %b %Y")

        # Open log file
        with open(log_file, "a") as log:
            log.write(f"\n{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Starting RV form generation for project: {project}\n")

            # Get the total number of rows to process
            sheet1 = wb[wb.sheetnames[0]]  # Parent data sheet (Sheet 1)
            total_rows = sum(1 for _ in sheet1.iter_rows(min_row=start_row, values_only=True) if _[1])
            progress_step = 100 / total_rows if total_rows > 0 else 100
            progress = 0

            # Loop through rows in Sheet 1 starting at the user-defined row
            for index, row in enumerate(sheet1.iter_rows(min_row=start_row, values_only=True), start=1):
                # Get the data from the row based on template type
                if template_type == "Instrument":
                    instrument_tag = format_value(row[1])  # Column B
                    manufacturer = format_value(row[4])   # Column E
                    model = format_value(row[5])          # Column F
                    process_connection = format_value(row[6])  # Column G
                    immersion_length = format_value(row[10])  # Column K
                    control_signal = format_value(row[18])    # Column S
                    min_range = format_value(row[13])         # Column N
                    max_range = format_value(row[14])         # Column O
                    unit = format_value(row[15])              # Column P
                    order_code = format_value(row[9])         # Column J (Order Code)

                    # Populate dynamic fields for the Instrument Template
                    new_sheet = wb.copy_worksheet(template_sheet)
                    rv_form_name = f"RV{str(index).zfill(2)}"
                    new_sheet.title = rv_form_name

                    new_sheet["A11"] = instrument_tag
                    new_sheet["C11"] = manufacturer
                    new_sheet["E11"] = model
                    new_sheet["A14"] = process_connection
                    new_sheet["B14"] = immersion_length
                    new_sheet["D14"] = control_signal
                    new_sheet["F14"] = min_range
                    new_sheet["G14"] = max_range
                    new_sheet["H14"] = unit
                    new_sheet["G11"] = order_code
                    new_sheet["I14"] = "N/A"
                else:
                    valve_tag = format_value(row[4])              # Column E (Valve Tag)
                    valve_make_model = format_value(row[6])       # Column G (Valve Make / Model Number)
                    actuator_make_model = format_value(row[7])    # Column H (Actuator Make / Model Number)
                    process_connection = format_value(row[9])     # Column J (Process Connection)
                    line_size = format_value(row[14])             # Column O (Line Size)
                    control_signal = format_value(row[16])        # Column Q (Control Signal)
                    dial_setting = format_value(row[13])          # Column N (Dial Setting)
                    flow_rate = format_value(row[11])             # Column L (Flow Rate)

                    # Populate dynamic fields for the Valve Template
                    new_sheet = wb.copy_worksheet(template_sheet)
                    rv_form_name = f"RV{str(index).zfill(2)}"
                    new_sheet.title = rv_form_name

                    new_sheet["A11"] = valve_tag
                    new_sheet["C11"] = valve_make_model
                    new_sheet["F11"] = actuator_make_model
                    new_sheet["A15"] = process_connection
                    new_sheet["B15"] = line_size
                    new_sheet["D13"] = control_signal
                    new_sheet["F15"] = dial_setting
                    new_sheet["I15"] = flow_rate

                # Populate static fields for both templates
                new_sheet["A5"] = project
                new_sheet["E5"] = client
                new_sheet["A7"] = reference_document
                new_sheet["E7"] = document_revision
                new_sheet["I5"] = current_date
                new_sheet["I7"] = rv_form_name

                # Apply alignment and font size to all populated cells
                font = Font(size=10)  # Slightly smaller font size
                for cell_ref in ["A5", "E5", "A7", "E7", "I5", "I7", "A11", "C11", "E11", "A14", "B14", "D14", "F14", "G14", "H14", "G11", "I14", "F11", "A15", "B15", "D13", "F15", "I15"]:
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

def auto_detect_start_row(sheet):
    """Automatically detect the starting row for instruments based on the Instrument Tag column."""
    for row_idx, row in enumerate(sheet.iter_rows(min_row=1, max_col=2, values_only=True), start=1):
        if row[1]:  # Check if the second column (Instrument Tag) has a value
            return row_idx
    return 6  # Default to row 6 if no valid row is found

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
        start_row = start_row_var.get()
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
            generate_rv_forms(input_file, output_file, project, client, reference_document, document_revision, int(start_row), template_type, progress_var, log_file)
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
    logo_image = PhotoImage(file="/mnt/data/subnetlogo.png")  # Adjust the path as needed
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


