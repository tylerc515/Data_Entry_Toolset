import os
import openpyxl
from tkinter import Tk, filedialog, messagebox, Button, Label, Checkbutton, IntVar, StringVar, ttk
import pandas as pd
import csv

def convert_all_csv_to_xlsx(parent_folder):
    """Converts all .csv files in the folder to .xlsx before processing."""
    print(f"Scanning for .csv files in folder: {parent_folder}")
    for root, _, files in os.walk(parent_folder):
        for file in files:
            if file.lower().endswith('.csv'):
                csv_path = os.path.join(root, file)
                try:
                    # Read the raw CSV and handle irregularities
                    with open(csv_path, 'r', encoding='utf-8') as f:
                        reader = csv.reader(f)
                        rows = list(reader)

                    # Determine the maximum column count and normalize rows
                    max_columns = max(len(row) for row in rows)
                    normalized_rows = [row + [''] * (max_columns - len(row)) for row in rows]

                    # Convert to a DataFrame
                    df = pd.DataFrame(normalized_rows)

                    # Save directly to XLSX, preserving all rows and columns
                    xlsx_path = os.path.splitext(csv_path)[0] + '.xlsx'
                    df.to_excel(xlsx_path, index=False, header=False, engine="openpyxl")
                    print(f"Converted {csv_path} to {xlsx_path}")
                except Exception as e:
                    print(f"Error converting {csv_path} to .xlsx: {e}")

def process_files(parent_folder, remove_vs, remove_empty):
    processed_files = []
    issue_files = []

    print(f"Scanning for .xlsx files in folder: {parent_folder}")
    all_files = [
        os.path.join(root, file)
        for root, _, files in os.walk(parent_folder)
        for file in files if file.lower().endswith(('.xls', '.xlsx')) and not file.endswith('_processed.xlsx')
    ]
    print(f"Found {len(all_files)} files to process.")

    total_files = len(all_files)
    progress = 0

    for file_path in all_files:
        try:
            current_file_label.set(f"Processing: {file_path}")
            print(f"Processing file: {file_path}")
            
            df = pd.read_excel(file_path, engine="openpyxl", header=None)

            if remove_vs:
                df = df.apply(lambda col: col.map(lambda x: str(x).replace("V", "") if isinstance(x, str) and x.endswith("V") else x))

            if remove_empty:
                df.dropna(how="all", inplace=True)

            output_file = f"_{os.path.basename(file_path).split('.')[0]}_processed.xlsx"
            output_path = os.path.join(os.path.dirname(file_path), output_file)
            df.to_excel(output_path, index=False, header=False, engine="openpyxl")
            processed_files.append(output_path)
            print(f"Processed and saved file: {output_path}")
        except Exception as e:
            issue_files.append((file_path, str(e)))
            print(f"Error processing {file_path}: {e}")
        
        # Update progress bar
        progress += 1
        progress_var.set((progress / total_files) * 100)
        progress_bar.update()

    completion_message(processed_files, issue_files)

def completion_message(processed_files, issue_files):
    if issue_files:
        message = f"Processed files: {len(processed_files)}\nIssues encountered:\n" + "\n".join([f"{fp}: {err}" for fp, err in issue_files])
    else:
        message = f"All files processed successfully.\nProcessed files: {len(processed_files)}"
    print(message)
    messagebox.showinfo("Processing Complete", message)

def run_script():
    if not (remove_vs_var.get() or remove_empty_var.get()):
        print("No options selected. Exiting.")
        return

    # Construct a detailed warning message
    message = "Have you backed up your files?\n\nThis script will perform the following operations:\n\n"

    if remove_vs_var.get():
        message += "- Remove 'V' from any number followed by 'V'.\n"

    if remove_empty_var.get():
        message += "- Hide rows where all data cells in the specified range are empty.\n"

    message += "\nDo you want to continue?"

    if not messagebox.askyesno("Confirm", message):
        print("User canceled.")
        return

    folder = filedialog.askdirectory(title="Select Parent Folder")
    if not folder:
        print("No folder selected. Exiting.")
        return

    print(f"Selected folder: {folder}")

    # Convert all CSV files to XLSX first
    convert_all_csv_to_xlsx(folder)

    # Check for already processed files
    processed_files_detected = [
        os.path.join(root, file)
        for root, _, files in os.walk(folder)
        for file in files if file.endswith('_processed.xlsx')
    ]

    if processed_files_detected:
        detected_folders = "\n".join(set(os.path.dirname(fp) for fp in processed_files_detected))
        warning_message = f"Processed files detected in the following folders:\n\n{detected_folders}\n\nDo you want to continue?"
        if not messagebox.askyesno("Warning", warning_message):
            print("User chose not to proceed with already processed files.")
            return

    process_files(folder, remove_vs_var.get(), remove_empty_var.get())

def create_gui():
    root = Tk()
    root.title("JTC Data Entry Toolkit")

    # Center the window dynamically
    window_width = 500
    window_height = 400
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x_position = (screen_width // 2) - (window_width // 2)
    y_position = (screen_height // 2) - (window_height // 2)
    root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

    Label(root, text="JTC Data Entry Toolkit", font=("Arial", 16, "bold")).pack(pady=10)
    Label(root, text="This tool processes Excel and CSV files to remove 'V' from data and hide empty rows.", wraplength=400, justify="center").pack()

    Label(root, text="Warning: Please back up your files before proceeding!", font=("Arial", 10, "bold"), fg="red", wraplength=400, justify="center").pack(pady=5)

    global remove_vs_var, remove_empty_var
    remove_vs_var = IntVar()
    remove_empty_var = IntVar()

    Checkbutton(root, text="Remove Vs from Data", variable=remove_vs_var, command=update_run_button).pack(anchor="w", padx=20)
    Checkbutton(root, text="Remove Empty Data Rows", variable=remove_empty_var, command=update_run_button).pack(anchor="w", padx=20)

    global current_file_label
    current_file_label = StringVar()
    current_file_label.set("")

    Label(root, textvariable=current_file_label, wraplength=480, justify="center").pack(pady=5)

    global progress_var
    progress_var = IntVar()
    global progress_bar
    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
    progress_bar.pack(fill="x", padx=20, pady=10)

    run_button = Button(root, text="Run", command=run_script, state="disabled", width=15)
    run_button.pack(pady=10)
    Button(root, text="Cancel", command=root.destroy, width=15).pack(pady=5)

    global run_button_global
    run_button_global = run_button

    root.mainloop()

def update_run_button():
    if remove_vs_var.get() or remove_empty_var.get():
        run_button_global.config(state="normal")
    else:
        run_button_global.config(state="disabled")

# Start the GUI
create_gui()
