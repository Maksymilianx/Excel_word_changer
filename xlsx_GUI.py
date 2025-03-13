import openpyxl
import os
import re
import webbrowser
from tkinter import Tk, Label, Entry, Button, filedialog, Text, Scrollbar, VERTICAL, END, IntVar, Checkbutton, \
    messagebox, Toplevel
import requests

GITHUB_REPO = "Maksymilianx/Excel_word_changer"
FALLBACK_VERSION = "1.0.0"  # Used if we can't fetch from GitHub


def fetch_latest_version():
    """Fetch the latest release tag from GitHub."""
    url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
    try:
        response = requests.get(url, timeout=5)
        response.raise_for_status()
        latest_version = response.json().get("tag_name", FALLBACK_VERSION)
        return latest_version
    except requests.RequestException:
        return FALLBACK_VERSION  # Fallback if GitHub is unreachable


VERSION = fetch_latest_version()  # Get the latest version


def remove_key_value_pair_from_cell(cell_value, key):
    """
    Remove the key-value pair and the surrounding pipes from the cell value.
    If the key is not found, return None.
    """
    pattern = r'\|?' + re.escape(key) + r'=[^|]*\|?'
    new_value = re.sub(pattern, '|', cell_value)
    new_value = re.sub(r'\|{2,}', '|', new_value).strip('|')
    return new_value if new_value != cell_value else None


def search_replace_or_remove_key(file_path, key, new_value, remove_key, log_widget, key_found):
    """Search for a key in an Excel file, replace its value, or remove it entirely."""
    try:
        workbook = openpyxl.load_workbook(file_path)
        found = False

        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        if remove_key:
                            updated_cell = remove_key_value_pair_from_cell(cell.value, key)
                        else:
                            pattern = rf"\|?{key}=[^|]*\|?"
                            updated_cell = re.sub(pattern, f"|{key}={new_value}|", cell.value)

                        if updated_cell is not None:
                            found = True
                            key_found[0] = True
                            cell.value = updated_cell

        if found:
            workbook.save(file_path)
            log_widget.insert(END, f"✅ Updated: {file_path}\n", "success")

    except Exception as e:
        log_widget.insert(END, f"❌ Error processing {file_path}: {e}\n", "error")


def process_excel_files(directory_path, key, new_value, remove_key, log_widget):
    """Process all Excel files in a directory."""
    log_widget.delete(1.0, END)
    log_widget.insert(END, f"🔄 Processing files in {directory_path}...\n", "info")

    key_found = [False]  # Use a list to track key presence across files

    try:
        for root, _, files in os.walk(directory_path):
            for file_name in files:
                if file_name.endswith('.xlsx'):
                    file_path = os.path.join(root, file_name)
                    search_replace_or_remove_key(file_path, key, new_value, remove_key, log_widget, key_found)

        if not key_found[0]:  # If the key was never found in any file
            log_widget.insert(END, f"⚠ Warning: The key '{key}' was not found in any file.\n", "warning")
            show_custom_warning_popup(f"The key '{key}' was not found in any file.")

        log_widget.insert(END, "✅ Process completed!\n", "success")

    except Exception as e:
        log_widget.insert(END, f"❌ An error occurred: {e}\n", "error")


def show_custom_warning_popup(message):
    """Custom popup for warnings, centered on the screen."""
    popup = Toplevel()
    popup.title("Warning")
    popup.geometry("300x100")

    popup.update_idletasks()
    screen_width = popup.winfo_screenwidth()
    screen_height = popup.winfo_screenheight()
    x_position = (screen_width // 2) - (300 // 2)
    y_position = (screen_height // 2) - (100 // 2)

    popup.geometry(f"300x100+{x_position}+{y_position}")  # Center the popup

    label = Label(popup, text=message, fg="red", font=("Arial", 12, "bold"))
    label.pack(pady=10)

    close_button = Button(popup, text="OK", command=popup.destroy)
    close_button.pack(pady=5)

    popup.grab_set()  # Make the popup modal (forces user to close it first)


def check_for_updates():
    """Check if a newer version is available and show a popup."""
    latest_version = fetch_latest_version()
    if latest_version > VERSION:  # Compare versions
        messagebox.showinfo("Update Available",
                            f"A new version ({latest_version}) is available!\nVisit GitHub to download.")
    else:
        messagebox.showinfo("Up to Date", "You are using the latest version.")


def browse_directory(entry_widget):
    """Open a directory selection dialog."""
    directory = filedialog.askdirectory()
    if directory:
        entry_widget.delete(0, END)
        entry_widget.insert(0, directory)


def start_processing(directory_entry, key_entry, value_entry, remove_key_var, log_widget):
    """Validate input and start processing files."""
    directory = directory_entry.get()
    key = key_entry.get()
    new_value = value_entry.get()
    remove_key = remove_key_var.get()

    if not directory or not os.path.exists(directory):
        log_widget.insert(END, "❌ Please select a valid directory.\n", "error")
        return

    if not key:
        log_widget.insert(END, "❌ Please enter a key to search for.\n", "error")
        return

    if not remove_key and not new_value:
        log_widget.insert(END, "❌ Please enter a new value or check the 'Remove Key' option.\n", "error")
        return

    if remove_key:
        confirm = messagebox.askyesno("Confirm Deletion", f"Are you sure you want to remove '{key}' and its value?")
        if not confirm:
            log_widget.insert(END, "⚠ Deletion canceled by user.\n", "warning")
            return

    process_excel_files(directory, key, new_value, remove_key, log_widget)


def open_github_link():
    """Opens the GitHub link in the default web browser."""
    webbrowser.open("https://github.com/Maksymilianx/Excel_word_changer")


# GUI Setup
root = Tk()
root.title(f"Excel Key-Value Updater - v{VERSION}")

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

# Remove Key Checkbox
remove_key_var = IntVar()
remove_key_checkbox = Checkbutton(root, text="Remove Key", variable=remove_key_var)
remove_key_checkbox.grid(row=3, column=1, pady=5)

# Start button
Button(root, text="Start Processing",
       command=lambda: start_processing(directory_entry, key_entry, value_entry, remove_key_var, log_widget)).grid(
    row=4, column=1, pady=10)

# Log display
log_widget = Text(root, height=15, width=70)
log_widget.grid(row=5, column=0, columnspan=3, padx=10, pady=10)
scrollbar = Scrollbar(root, orient=VERTICAL, command=log_widget.yview)
scrollbar.grid(row=5, column=3, sticky="ns")
log_widget.config(yscrollcommand=scrollbar.set)

# Message colors
log_widget.tag_config("error", foreground="red")
log_widget.tag_config("warning", foreground="orange")
log_widget.tag_config("success", foreground="green")
log_widget.tag_config("info", foreground="blue")

# Github sync updates
Button(root, text="Check for Updates", command=check_for_updates).grid(row=6, column=1, pady=5)

# GitHub link
github_label = Label(root, text="View on GitHub", fg="blue", cursor="hand2")
github_label.grid(row=7, column=1, pady=10)
github_label.bind("<Button-1>", lambda e: open_github_link())

log_widget.insert(END, f"🛠 Version: {VERSION}\n", "info")

root.mainloop()
