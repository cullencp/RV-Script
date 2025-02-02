import openpyxl
from openpyxl.styles import Alignment, Font
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os
from tkinter import PhotoImage

# Constants for cell references and formatting
HEADER_KEYWORDS = ["tag", "manufacturer", "model", "process connection", "immersion length", 
                   "control signal", "min range", "max range", "unit", "order code", 
                   "valve make / model number", "actuator make / model number", 
                   "line size", "dial setting", "flow rate"]
DEFAULT_START_ROW = 6
FONT_SIZE = 11

def get_sheet_by_partial_name(wb, partial_name):
    """Find a sheet by partial name match."""
    for sheet in wb.sheetnames:
        if partial_name.lower() in sheet.lower():
            return sheet
    raise ValueError(f"Sheet with partial name '{partial_name}' not found.")

def format_value(value):
    """Format the value in uppercase or return 'N/A' if empty."""
    if value is None or (isinstance(value, str) and value.strip().lower() == "n/a"):
        return "N/A"
    if isinstance(value, str):
        return value.upper()
    return value

def detect_header_row(sheet):
    """Dynamically detect the header row based on keyword matching."""
    for row_idx, row in enumerate(sheet.iter_rows(min_row=1, max_row=10, values_only=True), start=1):
        if any(any(keyword in str(cell).lower() for keyword in HEADER_KEYWORDS) for cell in row):
            return row_idx
    raise ValueError("Header row not found in the first 10 rows.")

def get_column_indices(sheet, header_row):
    """Get a dictionary mapping header names to column indices."""
    header_row_data = next(sheet.iter_rows(min_row=header_row, max_row=header_row, values_only=True))
    print("Detected Header Row:", header_row_data)  # Debug statement
    column_indices = {}
    for idx, header in enumerate(header_row_data):
        if header:  # Skip empty headers
            column_indices[header.strip().lower()] = idx
    return column_indices

def map_headers_to_required_fields(column_indices, required_fields):
    """Map detected headers to required fields using flexible matching."""
    header_mapping = {}
    for field, possible_names in required_fields.items():
        for name in possible_names:
            for header in column_indices:
                if name.lower() in header.lower():  # Case-insensitive partial match
                    header_mapping[field] = column_indices[header]
                    break
            else:
                continue
            break
        else:
            print(f"Warning: Required field '{field}' not found in the input file. Possible names: {possible_names}")
            header_mapping[field] = None  # Mark the field as missing
    return header_mapping

def populate_instrument_template(new_sheet, row_data, header_mapping):
    """Populate the Instrument template with data from the row.
       If the tag cell contains multiple lines, only the first line is used."""
    tag_val = row_data[header_mapping["tag"]]
    if tag_val is not None:
        tag_val = tag_val.splitlines()[0].strip()
    new_sheet["A11"] = format_value(tag_val)
    new_sheet["C11"] = format_value(row_data[header_mapping["manufacturer"]])
    new_sheet["E11"] = format_value(row_data[header_mapping["model"]])
    new_sheet["A14"] = format_value(row_data[header_mapping["process connection"]])
    new_sheet["B14"] = format_value(row_data[header_mapping["immersion length"]])
    new_sheet["D14"] = format_value(row_data[header_mapping["control signal"]])
    new_sheet["F14"] = format_value(row_data[header_mapping["min range"]])
    new_sheet["G14"] = format_value(row_data[header_mapping["max range"]])
    new_sheet["H14"] = format_value(row_data[header_mapping["unit"]])
    new_sheet["G11"] = format_value(row_data[header_mapping["order code"]])
    new_sheet["I14"] = "N/A"

def populate_valve_template(new_sheet, row_data, header_mapping):
    """Populate the Valve template with data from the row.
       Only the first tag number is used if multiple exist."""
    tag_val = row_data[header_mapping["tag"]]
    if tag_val is not None:
        tag_val = tag_val.splitlines()[0].strip()
    new_sheet["A11"] = format_value(tag_val)
    new_sheet["C11"] = format_value(row_data[header_mapping["valve make / model number"]])
    new_sheet["F11"] = format_value(row_data[header_mapping["actuator make / model number"]])
    new_sheet["A15"] = format_value(row_data[header_mapping["process connection"]])
    new_sheet["B15"] = format_value(row_data[header_mapping["line size"]])
    new_sheet["D13"] = format_value(row_data[header_mapping["control signal"]])
    new_sheet["F15"] = format_value(row_data[header_mapping["dial setting"]])
    new_sheet["I15"] = format_value(row_data[header_mapping["flow rate"]])

def apply_formatting(new_sheet):
    """Apply alignment and font size to all populated cells."""
    font = Font(size=FONT_SIZE)
    for cell_ref in ["A5", "E5", "A7", "E7", "I5", "I7", "A11", "C11", "E11", 
                     "A14", "B14", "D14", "F14", "G14", "H14", "G11", "I14", 
                     "F11", "A15", "B15", "D13", "F15", "I15"]:
        cell = new_sheet[cell_ref]
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.font = font

def generate_rv_forms(input_file, output_file, project, client, reference_document, 
                      document_revision, start_row, template_type, progress_var, log_file):
    try:
        print(f"Loading input file: {input_file}")
        wb = openpyxl.load_workbook(input_file)

        # Select the template sheet and required fields based on the template type
        if template_type == "Instrument":
            template_sheet_name = get_sheet_by_partial_name(wb, "RV Instrument  SUB-TF-01")
            required_fields = {
                "tag": ["Tag", "Instrument Tag", "BMS Tag"],
                "manufacturer": ["Manufacturer", "Instrument Manufacturer"],
                "model": ["model", "instrument model", "model number"],
                "process connection": ["process connection", "connection"],
                "immersion length": ["immersion length", "Immersion Length (mm)"],
                "control signal": ["control signal", "actuator control signal"],
                "min range": ["min range", "range min", "minimum range"],
                "max range": ["max range", "range max", "maximum range"],
                "unit": ["unit", "units"],
                "order code": ["order code", "code"],
            }
        else:
            template_sheet_name = get_sheet_by_partial_name(wb, "RV Valve SUB-TF-02")
            required_fields = {
                "tag": ["Tag", "Instrument Tag", "BMS Tag"],
                "valve make / model number": ["valve make", "valve model", "valve make / model number"],
                "actuator make / model number": ["actuator make", "actuator model", "actuator make / model number"],
                "process connection": ["process connection", "connection"],
                "line size": ["line size", "Line Size (mm)"],
                "control signal": ["control signal", "actuator control signal"],
                "dial setting": ["dial setting", "setting"],
                "flow rate": ["flow rate", "rate"],
            }

        template_sheet = wb[template_sheet_name]
        current_date = datetime.now().strftime("%d %b %Y").upper()

        with open(log_file, "a") as log:
            log.write(f"\n{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Starting RV form generation for project: {project}\n")

            # Use the first sheet as the parent data sheet
            sheet1 = wb[wb.sheetnames[0]]

            # Attempt auto-detection of header row; use fallback if provided
            try:
                header_row = detect_header_row(sheet1)
                log.write(f"Detected header row: {header_row}\n")
            except ValueError as e:
                if start_row:
                    header_row = int(start_row)
                    log.write(f"Auto-detect failed. Using fallback header row: {header_row}\n")
                else:
                    raise e

            # Build header mapping
            column_indices = get_column_indices(sheet1, header_row)
            header_mapping = map_headers_to_required_fields(column_indices, required_fields)

            # Check that the required "tag" field exists
            check_field = "tag"
            if header_mapping.get(check_field) is None:
                raise ValueError(f"Required field '{check_field}' not found. Possible names: {required_fields[check_field]}")

            # Count total rows using the check field (for progress purposes)
            total_rows = sum(1 for _ in sheet1.iter_rows(min_row=header_row + 1, values_only=True) if _[header_mapping[check_field]])
            progress_step = 100 / total_rows if total_rows > 0 else 100
            progress = 0
            progress_var.set(0)

            # Use a set to track processed instrument IDs (from the "No." column)
            processed_ids = set()
            # Use a separate counter for naming RV forms sequentially
            rv_counter = 1

            for row in sheet1.iter_rows(min_row=header_row + 1, values_only=True):
                # Check that the tag field exists; if not, skip this row.
                tag_val = row[header_mapping[check_field]]
                if tag_val is None:
                    print("Skipping row because the tag field is missing.")
                    continue

                # Duplicate check: assume the "No." column (lowercase "no.") holds a unique instrument number.
                instrument_number = row[column_indices.get("no.", 0)]
                # If the instrument number is missing (None or empty), skip the row as a duplicate entry.
                if not instrument_number:
                    print("Skipping row due to empty instrument number (likely an extra tag row).")
                    continue
                if instrument_number in processed_ids:
                    print(f"Skipping duplicate row for instrument number: {instrument_number}")
                    continue
                processed_ids.add(instrument_number)

                try:
                    new_sheet = wb.copy_worksheet(template_sheet)
                    # Use the rv_counter to generate a sequential sheet name.
                    rv_form_name = f"RV{str(rv_counter).zfill(2)}"
                    new_sheet.title = rv_form_name

                    # Populate the appropriate template
                    if template_type == "Instrument":
                        populate_instrument_template(new_sheet, row, header_mapping)
                    else:
                        populate_valve_template(new_sheet, row, header_mapping)

                    # Populate static fields
                    new_sheet["A5"] = project
                    new_sheet["E5"] = client
                    new_sheet["A7"] = reference_document
                    new_sheet["E7"] = document_revision
                    new_sheet["I5"] = current_date
                    new_sheet["I7"] = rv_form_name

                    apply_formatting(new_sheet)
                    log.write(f"Processed: {rv_form_name} (Instrument Number: {instrument_number})\n")

                    # Increment the form counter only when a form is generated.
                    rv_counter += 1

                except Exception as e:
                    log.write(f"Error processing a row: {str(e)}\n")

                progress += progress_step
                progress_var.set(progress)

            # Hide the template sheet so that only the generated RV forms appear
            template_sheet.sheet_state = 'hidden'

            try:
                print(f"Saving output file: {output_file}")
                wb.save(output_file)
                print("File saved successfully.")
            except Exception as e:
                raise Exception(f"Failed to save the file: {str(e)}")

            progress_var.set(100)
            log.write("RV form generation completed successfully.\n")
            messagebox.showinfo("Success", f"RV forms generated successfully! Output saved as {output_file}")

    except Exception as e:
        with open(log_file, "a") as log:
            log.write(f"Error: {str(e)}\n")
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

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
        start_row = start_row_var.get().strip()  # Fallback header row (optional)
        output_folder = output_folder_var.get()
        template_type = template_type_var.get()

        if not input_file or not project or not client or not reference_document or not document_revision or not output_folder or not template_type:
            messagebox.showerror("Input Error", "Please fill in all fields and select files.")
            return

        file_extension = os.path.splitext(input_file)[-1]
        output_file = os.path.join(output_folder, f"{project}-RVs{file_extension}")
        log_file = os.path.join(output_folder, "rv_generator_log.txt")

        try:
            generate_rv_forms(input_file, output_file, project, client, reference_document,
                              document_revision, start_row, template_type, progress_var, log_file)
        except Exception as e:
            with open(log_file, "a") as log:
                log.write(f"Error: {str(e)}\n")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    root = tk.Tk()
    root.title("RV Form Generator")
    root.configure(bg="#ffffff")
    root.geometry("600x500")

    # Set the window icon using a relative path to the ICO file
    try:
        icon_path = os.path.join("resources", "subnetlogo.ico")
        root.iconbitmap(icon_path)
    except Exception as e:
        print("Window icon not set:", e)

    # Variables
    input_file_var = tk.StringVar()
    output_folder_var = tk.StringVar()
    project_var = tk.StringVar()
    client_var = tk.StringVar()
    reference_var = tk.StringVar()
    revision_var = tk.StringVar()
    start_row_var = tk.StringVar(value="")  # Optional fallback header row
    template_type_var = tk.StringVar(value="Instrument")
    progress_var = tk.DoubleVar()

    # Load the PNG logo using a relative path; ensure the file is located in the "resources" folder
    try:
        logo_path = os.path.join("resources", "subnetlogo.png")
        logo_image = PhotoImage(file=logo_path)
        logo_label = tk.Label(root, image=logo_image, bg="#ffffff")
        logo_label.grid(row=0, column=0, columnspan=3, pady=(10, 20))
    except Exception as e:
        print("Logo image not found:", e)

    pad_options = {'padx': 10, 'pady': 5}
    tk.Label(root, text="Input File:", bg="#ffffff", font=("Arial", 10)).grid(row=1, column=0, sticky="e", **pad_options)
    tk.Entry(root, textvariable=input_file_var, width=50).grid(row=1, column=1, **pad_options)
    tk.Button(root, text="Browse", command=browse_input_file, bg="#4CAF50", fg="#ffffff", font=("Arial", 10)).grid(row=1, column=2, **pad_options)

    tk.Label(root, text="Output Location:", bg="#ffffff", font=("Arial", 10)).grid(row=2, column=0, sticky="e", **pad_options)
    tk.Entry(root, textvariable=output_folder_var, width=50).grid(row=2, column=1, **pad_options)
    tk.Button(root, text="Browse", command=browse_output_location, bg="#4CAF50", fg="#ffffff", font=("Arial", 10)).grid(row=2, column=2, **pad_options)

    tk.Label(root, text="Template Type:", bg="#ffffff", font=("Arial", 10)).grid(row=3, column=0, sticky="e", **pad_options)
    template_dropdown = ttk.Combobox(root, textvariable=template_type_var, values=["Instrument", "Valve"], state="readonly", width=47)
    template_dropdown.grid(row=3, column=1, columnspan=2, **pad_options)

    tk.Label(root, text="Project Name:", bg="#ffffff", font=("Arial", 10)).grid(row=4, column=0, sticky="e", **pad_options)
    tk.Entry(root, textvariable=project_var, width=50).grid(row=4, column=1, columnspan=2, **pad_options)

    tk.Label(root, text="Client Name:", bg="#ffffff", font=("Arial", 10)).grid(row=5, column=0, sticky="e", **pad_options)
    tk.Entry(root, textvariable=client_var, width=50).grid(row=5, column=1, columnspan=2, **pad_options)

    tk.Label(root, text="Reference Document:", bg="#ffffff", font=("Arial", 10)).grid(row=6, column=0, sticky="e", **pad_options)
    tk.Entry(root, textvariable=reference_var, width=50).grid(row=6, column=1, columnspan=2, **pad_options)

    tk.Label(root, text="Document Revision:", bg="#ffffff", font=("Arial", 10)).grid(row=7, column=0, sticky="e", **pad_options)
    tk.Entry(root, textvariable=revision_var, width=50).grid(row=7, column=1, columnspan=2, **pad_options)

    tk.Label(root, text="Header Row Fallback (optional):", bg="#ffffff", font=("Arial", 10)).grid(row=8, column=0, sticky="e", **pad_options)
    tk.Entry(root, textvariable=start_row_var, width=50).grid(row=8, column=1, columnspan=2, **pad_options)

    tk.Button(root, text="Generate RV Forms", command=generate_forms, bg="#4CAF50", fg="#ffffff", font=("Arial", 12, "bold")).grid(row=9, column=0, columnspan=3, pady=15)

    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
    progress_bar.grid(row=10, column=0, columnspan=3, sticky="ew", padx=10, pady=10)

    tk.Label(root, text="© 2025 Subnet Ltd. All rights reserved.", bg="#ffffff", font=("Arial", 9)).grid(row=11, column=0, columnspan=3, pady=(10, 0))

    root.mainloop()

if __name__ == "__main__":
    main()

