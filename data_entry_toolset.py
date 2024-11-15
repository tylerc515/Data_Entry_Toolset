import os
import openpyxl
from tkinter import Tk, filedialog, messagebox, Button, Label, Checkbutton, IntVar, StringVar, ttk
import pandas as pd

def process_files(parent_folder, remove_vs, remove_empty):
    processed_files = []
    issue_files = []

    # Count only unprocessed files for progress
    all_files = [
        os.path.join(root, file)
        for root, _, files in os.walk(parent_folder)
        for file in files if file.endswith(('.xls', '.xlsx')) and not file.endswith('_processed.xlsx')
    ]
    total_files = len(all_files)
    progress = 0

    for file_path in all_files:
        try:
            current_file_label.set(f"Processing: {file_path}")
            if not file_path.endswith('.xlsx'):
                file_path = convert_to_xlsx(file_path)

            workbook = openpyxl.load_workbook(file_path)
            sheet = workbook.active

            start_col = 3
            end_col = sheet.max_column

            if remove_vs:
                remove_v_from_numbers(sheet, start_col, end_col)

            if remove_empty:
                hide_empty_elevations(sheet, start_col, end_col)

            # Save processed file with underscore prefix
            output_file = f"_{os.path.basename(file_path).replace('.xlsx', '_processed.xlsx')}"
            output_path = os.path.join(os.path.dirname(file_path), output_file)
            workbook.save(output_path)
            processed_files.append(output_path)
        except Exception as e:
            issue_files.append((file_path, str(e)))
        
        # Update progress bar
        progress += 1
        progress_var.set((progress / total_files) * 100)
        progress_bar.update()

    completion_message(processed_files, issue_files)

def remove_v_from_numbers(sheet, start_col, end_col):
    for row in sheet.iter_rows(min_row=2, min_col=start_col, max_col=end_col):
        for cell in row:
            if cell.value and isinstance(cell.value, str) and cell.value.endswith('V'):
                try:
                    numeric_part = float(cell.value[:-1])
                    cell.value = numeric_part
                except ValueError:
                    pass

def hide_empty_elevations(sheet, start_col, end_col):
    row = 2
    while row <= sheet.max_row:
        group_rows = [row, row + 1, row + 2]
        if all(
            all(sheet.cell(row=r, column=c).value is None for c in range(start_col, end_col + 1))
            for r in group_rows
        ):
            for r in group_rows:
                sheet.row_dimensions[r].hidden = True
        row += 3

def convert_to_xlsx(file_path):
    new_path = os.path.splitext(file_path)[0] + '.xlsx'
    df = pd.read_excel(file_path)
    df.to_excel(new_path, index=False)
    return new_path

def completion_message(processed_files, issue_files):
    if issue_files:
        message = f"Processed files: {len(processed_files)}\nIssues encountered:\n" + "\n".join([f"{fp}: {err}" for fp, err in issue_files])
    else:
        message = f"All files processed successfully.\nProcessed files: {len(processed_files)}"
    messagebox.showinfo("Processing Complete", message)

def run_script():
    if not (remove_vs_var.get() or remove_empty_var.get()):
        return

    # Construct a detailed warning message
    message = "Have you backed up your files?\n\nThis script will perform the following operations:\n\n"

    if remove_vs_var.get():
        message += "- Remove 'V' from any number followed by 'V'.\n"

    if remove_empty_var.get():
        message += "- Hide rows where all data cells in the specified range are empty.\n"

    message += "\nDo you want to continue?"

    if not messagebox.askyesno("Confirm", message):
        return

    folder = filedialog.askdirectory(title="Select Parent Folder")
    if not folder:
        return

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
    Label(root, text="This tool processes Excel files to remove 'V' from data and hide empty rows.", wraplength=400, justify="center").pack()

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
