import openpyxl
import os
import re
import webbrowser
from tkinter import Tk, Label, Entry, Button, filedialog, Text, Scrollbar, VERTICAL, END

def search_and_replace_key_in_excel(file_path, key, new_value, log_widget):
    """
    Reads an Excel file, searches for a specific key-value pair delimited by '|',
    and replaces the value associated with the key while maintaining the format.

    Args:
        file_path (str): Path to the Excel file.
        key (str): The key to search for (e.g., "LocId").
        new_value (str): The new value to replace the current value with (e.g., "R00").
        log_widget (Text): Text widget to log messages.
    """
    try:
        workbook = openpyxl.load_workbook(file_path)
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str):  # Check if the cell contains a string
                        pattern = rf"\|{key}=.*?\|"
                        if re.search(pattern, cell.value):  # Check if the key exists in the cell
                            updated_cell = re.sub(pattern, f"|{key}={new_value}|", cell.value)
                            cell.value = updated_cell
        workbook.save(file_path)
        log_widget.insert(END, f"Updated: {file_path}\n")
    except Exception as e:
        log_widget.insert(END, f"Error processing {file_path}: {e}\n")

def process_excel_files_in_directory(directory_path, key, new_value, log_widget):
    """
    Iterates over all Excel files in a directory (including subfolders) and performs
    search-and-replace for key-value pairs.

    Args:
        directory_path (str): Path to the directory containing Excel files.
        key (str): The key to search for (e.g., "LocId").
        new_value (str): The new value to replace the key's value (e.g., "R00").
        log_widget (Text): Text widget to log messages.
    """
    log_widget.delete(1.0, END)  # Clear previous logs
    log_widget.insert(END, f"Processing files in {directory_path}...\n")
    try:
        for root, _, files in os.walk(directory_path):
            for file_name in files:
                if file_name.endswith('.xlsx'):
                    file_path = os.path.join(root, file_name)
                    search_and_replace_key_in_excel(file_path, key, new_value, log_widget)
        log_widget.insert(END, "Process completed!\n")
    except Exception as e:
        log_widget.insert(END, f"An error occurred: {e}\n")

def browse_directory(entry_widget):
    """
    Opens a file dialog to select a directory and updates the entry widget.
    """
    directory = filedialog.askdirectory()
    if directory:
        entry_widget.delete(0, END)
        entry_widget.insert(0, directory)

def start_processing(directory_entry, key_entry, value_entry, log_widget):
    """
    Starts the processing of Excel files using the provided input values.
    """
    directory = directory_entry.get()
    key = key_entry.get()
    new_value = value_entry.get()

    if not directory or not os.path.exists(directory):
        log_widget.insert(END, "Please select a valid directory.\n")
        return

    if not key:
        log_widget.insert(END, "Please enter a key to search for.\n")
        return

    if not new_value:
        log_widget.insert(END, "Please enter a new value to replace with.\n")
        return

    process_excel_files_in_directory(directory, key, new_value, log_widget)

def open_github_link():
    """
    Opens the GitHub link in the default web browser.
    """
    webbrowser.open("https://github.com/Maksymilianx/Excel_word_changer")

# GUI Setup
root = Tk()
root.title("Excel Key-Value Updater")

# Directory selection
Label(root, text="Select Directory:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
directory_entry = Entry(root, width=50)
directory_entry.grid(row=0, column=1, padx=10, pady=5)
browse_button = Button(root, text="Browse", command=lambda: browse_directory(directory_entry))
browse_button.grid(row=0, column=2, padx=10, pady=5)

# Key input
Label(root, text="Key to Search:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
key_entry = Entry(root, width=50)
key_entry.grid(row=1, column=1, padx=10, pady=5)

# New value input
Label(root, text="New Value:").grid(row=2, column=0, padx=10, pady=5, sticky="w")
value_entry = Entry(root, width=50)
value_entry.grid(row=2, column=1, padx=10, pady=5)

# Start button
start_button = Button(root, text="Start Processing", command=lambda: start_processing(directory_entry, key_entry, value_entry, log_widget))
start_button.grid(row=3, column=1, pady=10)

# Log display
log_widget = Text(root, height=15, width=70)
log_widget.grid(row=4, column=0, columnspan=3, padx=10, pady=10)
scrollbar = Scrollbar(root, orient=VERTICAL, command=log_widget.yview)
scrollbar.grid(row=4, column=3, sticky="ns")
log_widget.config(yscrollcommand=scrollbar.set)

# GitHub link
github_label = Label(root, text="View on GitHub", fg="blue", cursor="hand2")
github_label.grid(row=5, column=1, pady=10)
github_label.bind("<Button-1>", lambda e: open_github_link())

root.mainloop()
