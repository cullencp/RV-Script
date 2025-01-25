import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

def generate_rv_forms(input_file, output_file, project, client, reference_document, document_revision, start_row):
    try:
        # Load the workbook
        wb = openpyxl.load_workbook(input_file)

        # Get the sheets
        sheet1 = wb[wb.sheetnames[0]]  # Parent data sheet (Sheet 1)
        template_sheet = wb["RV Instrument  SUB-TF-01"]  # Template (Sheet 3)

        # Current date in the required format
        current_date = datetime.now().strftime("%d %b %Y")

        # Loop through rows in Sheet 1 starting at the user-defined row
        for index, row in enumerate(sheet1.iter_rows(min_row=start_row, values_only=True), start=1):
            # Function to handle values and ensure uppercase "N/A"
            def format_value(value):
                if value is None or (isinstance(value, str) and value.strip().lower() == "n/a"):
                    return "N/A"
                return str(value).upper() if isinstance(value, str) else value

            # Get the data from the row
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
            new_sheet["G11"] = order_code
            new_sheet["I14"] = "N/A"

            # Apply alignment to all populated cells
            for cell_ref in ["A5", "E5", "A7", "E7", "I5", "I7", "A11", "C11", "E11", "A14", "B14", "D14", "F14", "G14", "H14", "I14", "G11"]:
                cell = new_sheet[cell_ref]
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Save the updated workbook
        wb.save(output_file)
        messagebox.showinfo("Success", f"RV forms generated successfully! Output saved as {output_file}")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def auto_detect_start_row(sheet):
    """Automatically detect the starting row for instruments based on the Instrument Tag column."""
    for row in sheet.iter_rows(min_row=1, max_col=2, values_only=True):
        if row[1]:  # Check if the second column (Instrument Tag) has a value
            return row[0]
    return 6  # Default to row 6 if no valid row is found

# GUI Implementation
def main():
    def browse_input_file():
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        input_file_var.set(file_path)

    def browse_output_file():
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        output_file_var.set(file_path)

    def generate_forms():
        input_file = input_file_var.get()
        output_file = output_file_var.get()
        project = project_var.get()
        client = client_var.get()
        reference_document = reference_var.get()
        document_revision = revision_var.get()
        start_row = start_row_var.get()

        if not input_file or not output_file or not project or not client or not reference_document or not document_revision:
            messagebox.showerror("Input Error", "Please fill in all fields and select files.")
            return

        try:
            wb = openpyxl.load_workbook(input_file)
            sheet1 = wb[wb.sheetnames[0]]
            if not start_row:
                start_row = auto_detect_start_row(sheet1)

            generate_rv_forms(input_file, output_file, project, client, reference_document, document_revision, int(start_row))
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    root = tk.Tk()
    root.title("RV Form Generator")

    # Variables
    input_file_var = tk.StringVar()
    output_file_var = tk.StringVar()
    project_var = tk.StringVar()
    client_var = tk.StringVar()
    reference_var = tk.StringVar()
    revision_var = tk.StringVar()
    start_row_var = tk.StringVar()

    # GUI Layout
    tk.Label(root, text="Input File:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(root, textvariable=input_file_var, width=50).grid(row=0, column=1, padx=5, pady=5)
    tk.Button(root, text="Browse", command=browse_input_file).grid(row=0, column=2, padx=5, pady=5)

    tk.Label(root, text="Output File:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(root, textvariable=output_file_var, width=50).grid(row=1, column=1, padx=5, pady=5)
    tk.Button(root, text="Browse", command=browse_output_file).grid(row=1, column=2, padx=5, pady=5)

    tk.Label(root, text="Project Name:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(root, textvariable=project_var, width=50).grid(row=2, column=1, padx=5, pady=5)

    tk.Label(root, text="Client Name:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(root, textvariable=client_var, width=50).grid(row=3, column=1, padx=5, pady=5)

    tk.Label(root, text="Reference Document:").grid(row=4, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(root, textvariable=reference_var, width=50).grid(row=4, column=1, padx=5, pady=5)

    tk.Label(root, text="Document Revision:").grid(row=5, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(root, textvariable=revision_var, width=50).grid(row=5, column=1, padx=5, pady=5)

    tk.Label(root, text="Starting Row (Auto-Detect if Blank):").grid(row=6, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(root, textvariable=start_row_var, width=50).grid(row=6, column=1, padx=5, pady=5)

    tk.Button(root, text="Generate RV Forms", command=generate_forms).grid(row=7, column=0, columnspan=3, pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()

