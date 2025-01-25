import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os
from openpyxl.drawing.image import Image
from tkinter import PhotoImage

def generate_rv_forms(input_file, output_file, project, client, reference_document, document_revision, start_row, progress_var, log_file):
    try:
        # Load the workbook
        wb = openpyxl.load_workbook(input_file)

        # Get the sheets
        sheet1 = wb[wb.sheetnames[0]]  # Parent data sheet (Sheet 1)
        template_sheet = wb["RV Instrument  SUB-TF-01"]  # Template (Sheet 3)

        # Load the logo image
        logo = Image("C:\\Users\\PC\\Desktop\\RV-Script\\subnetlogo.png")  # Path to the company logo
        logo.anchor = "E1"  # Position the logo in the center top of the template

        # Current date in the required format
        current_date = datetime.now().strftime("%d %b %Y")

        # Open log file
        with open(log_file, "a") as log:
            log.write(f"\n{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Starting RV form generation for project: {project}\n")

            # Get the total number of rows to process
            total_rows = sum(1 for _ in sheet1.iter_rows(min_row=start_row, values_only=True) if _[1])
            progress_step = 100 / total_rows if total_rows > 0 else 100
            progress = 0

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

                # Insert the logo into the new sheet
                new_sheet.add_image(logo)

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

    def generate_forms():
        input_file = input_file_var.get()
        project = project_var.get()
        client = client_var.get()
        reference_document = reference_var.get()
        document_revision = revision_var.get()
        start_row = start_row_var.get()

        if not input_file or not project or not client or not reference_document or not document_revision:
            messagebox.showerror("Input Error", "Please fill in all fields and select files.")
            return

        # Generate default output file name based on project name
        file_extension = os.path.splitext(input_file)[-1]
        output_file = f"{project}-RVs{file_extension}"

        log_file = "rv_generator_log.txt"

        try:
            wb = openpyxl.load_workbook(input_file)
            sheet1 = wb[wb.sheetnames[0]]
            if not start_row:
                start_row = auto_detect_start_row(sheet1)

            generate_rv_forms(input_file, output_file, project, client, reference_document, document_revision, int(start_row), progress_var, log_file)
        except Exception as e:
            with open(log_file, "a") as log:
                log.write(f"Error: {str(e)}\n")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    root = tk.Tk()
    root.title("RV Form Generator")
    root.configure(bg="#00274d")  # Background color matching the logo

    # Variables
    input_file_var = tk.StringVar()
    project_var = tk.StringVar()
    client_var = tk.StringVar()
    reference_var = tk.StringVar()
    revision_var = tk.StringVar()
    start_row_var = tk.StringVar()
    progress_var = tk.DoubleVar()

    # Add the logo to the GUI
    logo_image = PhotoImage(file="C:\\Users\\PC\\Desktop\\RV-Script\\subnetlogo.png")  # Adjust the path as needed
    logo_label = tk.Label(root, image=logo_image, bg="#00274d")
    logo_label.grid(row=0, column=0, columnspan=3, pady=(10, 20))  # Adjust padding for spacing

    # GUI Layout
    tk.Label(root, text="Input File:", bg="#00274d", fg="white").grid(row=1, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(root, textvariable=input_file_var, width=50).grid(row=1, column=1, padx=5, pady=5)
    tk.Button(root, text="Browse", command=browse_input_file).grid(row=1, column=2, padx=5, pady=5)

    tk.Label(root, text="Project Name:", bg="#00274d", fg="white").grid(row=2, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(root, textvariable=project_var, width=50).grid(row=2, column=1, padx=5, pady=5)

    tk.Label(root, text="Client Name:", bg="#00274d", fg="white").grid(row=3, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(root, textvariable=client_var, width=50).grid(row=3, column=1, padx=5, pady=5)

    tk.Label(root, text="Reference Document:", bg="#00274d", fg="white").grid(row=4, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(root, textvariable=reference_var, width=50).grid(row=4, column=1, padx=5, pady=5)

    tk.Label(root, text="Document Revision:", bg="#00274d", fg="white").grid(row=5, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(root, textvariable=revision_var, width=50).grid(row=5, column=1, padx=5, pady=5)

    tk.Label(root, text="Starting Row (Auto-Detect if Blank):", bg="#00274d", fg="white").grid(row=6, column=0, padx=5, pady=5, sticky="e")
    tk.Entry(root, textvariable=start_row_var, width=50).grid(row=6, column=1, padx=5, pady=5)

    tk.Button(root, text="Generate RV Forms", command=generate_forms).grid(row=7, column=0, columnspan=3, pady=10)

    # Progress Bar
    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
    progress_bar.grid(row=8, column=0, columnspan=3, padx=5, pady=10, sticky="ew")

    root.mainloop()

if __name__ == "__main__":
    main()



