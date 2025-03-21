# (CC0) balaji.work

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import ifcopenshell
import csv
import os
import sys
import pandas as pd
import webbrowser
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# Log errors to a file when running as an exe
def log_errors_to_file(log_file="error_log.txt"):
    sys.stdout = open(log_file, "w")
    sys.stderr = sys.stdout

if getattr(sys, 'frozen', False):
    log_errors_to_file()

def extract_ifc_properties(ifc_file_path, output_csv_path):
    print(f"Processing: {ifc_file_path}")
    ifc_file = ifcopenshell.open(ifc_file_path)
    elements = ifc_file.by_type("IfcElement")

    property_names = set()
    element_data = []

    for element in elements:
        element_id = element.GlobalId
        element_name = element.Name if element.Name else "Unknown"
        element_type = element.is_a()
        properties = {"GlobalId": element_id, "Name": element_name, "Type": element_type}

        if hasattr(element, "IsDefinedBy"):
            for rel in element.IsDefinedBy:
                if rel.is_a("IfcRelDefinesByProperties"):
                    prop_set = rel.RelatingPropertyDefinition
                    if hasattr(prop_set, "HasProperties"):
                        for prop in prop_set.HasProperties:
                            if hasattr(prop, "Name") and hasattr(prop, "NominalValue"):
                                properties[prop.Name] = prop.NominalValue.wrappedValue
                                property_names.add(prop.Name)

        element_data.append(properties)

    property_names = ["GlobalId", "Name", "Type"] + sorted(property_names)
    with open(output_csv_path, mode="w", newline="", encoding="utf-8") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=property_names)
        writer.writeheader()
        for data in element_data:
            writer.writerow(data)
    print(f"Saved CSV: {output_csv_path}")

def process_ifc_directory(input_dir, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    
    files_to_process = [f for root, _, files in os.walk(input_dir) for f in files if f.lower().endswith(".ifc")]
    total_files = len(files_to_process)
    
    if total_files == 0:
        print("No IFC files found.")
        return

    file_count = 0
    for root, _, files in os.walk(input_dir):
        for file_name in files:
            if file_name.lower().endswith(".ifc"):
                file_count += 1
                input_file_path = os.path.join(root, file_name)
                relative_path = os.path.relpath(root, input_dir)
                output_subdir = os.path.join(output_dir, relative_path)
                os.makedirs(output_subdir, exist_ok=True)
                output_file_name = os.path.splitext(file_name)[0] + ".csv"
                output_file_path = os.path.join(output_subdir, output_file_name)
                
                # Update progress bar and terminal output
                progress_var.set((file_count / total_files) * 100)
                progress_label.config(text=f"Processing file {file_count} of {total_files}: {file_name}")
                app.update_idletasks()

                print(f"Processing file {file_count} of {total_files}: {file_name}")
                try:
                    extract_ifc_properties(input_file_path, output_file_path)
                except Exception as e:
                    print(f"Error processing {input_file_path}: {str(e)}")

def create_combined_excel(output_dir):
    combined_data = []
    all_columns = set()

    for root, _, files in os.walk(output_dir):
        for file_name in files:
            if file_name.lower().endswith(".csv"):
                csv_file_path = os.path.join(root, file_name)
                df = pd.read_csv(csv_file_path)
                all_columns.update(df.columns)
                combined_data.append(df)

    if combined_data:
        # Ensure all columns are retained and combined correctly
        final_df = pd.concat(combined_data, ignore_index=True, sort=False)
        final_df = final_df.reindex(columns=all_columns)  # Reorder columns to maintain consistency
    else:
        # Create an empty DataFrame if no valid data found
        print("No valid data found. Creating an empty validation file.")
        final_df = pd.DataFrame(columns=all_columns)

    validation_file = os.path.join(output_dir, "validation_output.xlsx")
    print(f"Saving validation file to: {validation_file}")

    try:
        final_df.to_excel(validation_file, index=False)
        validate_excel(validation_file)
        print(f"Validation file created at: {validation_file}")
    except Exception as e:
        print(f"Error creating Excel: {e}")

def validate_excel(excel_path):
    wb = load_workbook(excel_path)
    ws = wb.active
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    # Get header row and map column names to their positions
    header_row = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]

    # Dynamically locate target columns
    try:
        idx_CCILevel1ParentLocationID = header_row.index("CCILevel1ParentLocationID")
        idx_CCILevel1ParentTypeID = header_row.index("CCILevel1ParentTypeID")
        idx_CCILevel2ParentLocationID = header_row.index("CCILevel2ParentLocationID")
        idx_CCILevel2ParentTypeID = header_row.index("CCILevel2ParentTypeID")
        idx_CCILocationID = header_row.index("CCILocationID")
        idx_CCIMultiLevelLocationID = header_row.index("CCIMultiLevelLocationID")
        idx_CCIMultiLevelTypeID = header_row.index("CCIMultiLevelTypeID")
    except ValueError as e:
        print(f"Required columns not found in the sheet: {e}")
        return

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        # Get values from dynamically identified columns
        CCILevel2ParentTypeID = str(row[idx_CCILevel2ParentTypeID].value)
        CCILevel1ParentTypeID = str(row[idx_CCILevel1ParentTypeID].value)
        expected_type_id = f"§{CCILevel2ParentTypeID}.{CCILevel1ParentTypeID}"
        if str(row[idx_CCIMultiLevelTypeID].value) != expected_type_id:
            row[idx_CCIMultiLevelTypeID].fill = red_fill

        CCILevel2ParentLocationID = str(row[idx_CCILevel2ParentLocationID].value)
        CCILevel1ParentLocationID = str(row[idx_CCILevel1ParentLocationID].value)
        CCILocationID = str(row[idx_CCILocationID].value)
        expected_location_id = f"+{CCILevel2ParentLocationID}.{CCILevel1ParentLocationID}.{CCILocationID}"
        if str(row[idx_CCIMultiLevelLocationID].value) != expected_location_id:
            row[idx_CCIMultiLevelLocationID].fill = red_fill

    wb.save(excel_path)

def select_input_directory():
    input_dir = filedialog.askdirectory(title="Select Input Directory")
    input_dir_entry.delete(0, tk.END)
    input_dir_entry.insert(0, input_dir)

def select_output_directory():
    output_dir = filedialog.askdirectory(title="Select Output Directory")
    output_dir_entry.delete(0, tk.END)
    output_dir_entry.insert(0, output_dir)

def start_processing():
    input_dir = input_dir_entry.get()
    output_dir = output_dir_entry.get()
    
    if not input_dir or not output_dir:
        messagebox.showerror("Error", "Please select both input and output directories.")
        return

    try:
        print("Starting IFC to CSV conversion...")
        progress_var.set(0)
        progress_label.config(text="Initializing file processing...")
        process_ifc_directory(input_dir, output_dir)
        create_combined_excel(output_dir)
        messagebox.showinfo("Success", "Processing complete! CSV files and validation sheet saved.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Create Tkinter app
app = tk.Tk()
app.title("IFC to CSV Converter - (CC) balaji.work")
app.geometry("500x600")

tk.Label(app, text="Created for Eastern Ring Road Project by Balaji Balagurusami babs@cowi.com for COWI A/S.", fg="blue").pack(pady=5)
tk.Label(app, text="License: Creative Commons 0 1.0 Universal", fg="blue").pack(pady=5)

def open_github():
    webbrowser.open_new("https://github.com/balajibalagurusami/python/ifc2csv")

github_label = tk.Label(app, text="Source: GitHub Repository", fg="blue", cursor="hand2")
github_label.pack(pady=5)
github_label.bind("<Button-1>", lambda e: open_github())


# Input Directory
tk.Label(app, text="Input Directory:").pack(pady=5)
input_dir_entry = tk.Entry(app, width=50)
input_dir_entry.pack(pady=2)
tk.Button(app, text="Browse", command=select_input_directory).pack(pady=2)

# Output Directory
tk.Label(app, text="Output Directory:").pack(pady=5)
output_dir_entry = tk.Entry(app, width=50)
output_dir_entry.pack(pady=2)
tk.Button(app, text="Browse", command=select_output_directory).pack(pady=2)

# Start Button
tk.Button(app, text="Start Processing", command=start_processing, bg="green", fg="white").pack(pady=20)

# Progress Bar
progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(app, variable=progress_var, maximum=100)
progress_bar.pack(pady=5, fill=tk.X, padx=20)

# Progress Label
progress_label = tk.Label(app, text="Waiting for file selection...")
progress_label.pack(pady=5)

app.mainloop()
