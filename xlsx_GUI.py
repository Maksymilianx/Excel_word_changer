from tkinter import Tk, Frame, Label, Entry, Button, Text, Scrollbar, VERTICAL, END, IntVar, Checkbutton
from tkinter.ttk import Notebook, Progressbar
import threading
from functions import (
    VERSION,
    start_processing,
    start_value_replacement,
    check_for_updates,
    browse_directory,
    open_github_link
)
from tooltip import CreateToolTip


def launch_gui():
    root = Tk()
    root.title(f"xlxs fixer - v{VERSION}")

    notebook = Notebook(root)
    notebook.pack(expand=True, fill="both")

    # ----- Flat file fixer Tab (Processing Functionality) -----
    flat_tab = Frame(notebook)
    notebook.add(flat_tab, text="Flat file fixer")

    # Row 0: Processing Directory
    Label(flat_tab, text="?", bg="blue", fg="white", font=("Arial", 8, "bold")).grid(row=0, column=0, padx=2, pady=5,
                                                                                     sticky="e")
    Label(flat_tab, text="Select Processing Directory:").grid(row=0, column=1, padx=10, pady=5, sticky="w")
    directory_entry = Entry(flat_tab, width=40)
    directory_entry.grid(row=0, column=2, padx=10, pady=5)
    Button(flat_tab, text="Browse", command=lambda: browse_directory(directory_entry)).grid(row=0, column=3, padx=10,
                                                                                            pady=5)
    CreateToolTip(flat_tab.grid_slaves(row=0, column=0)[0], "Choose the folder containing the Excel files to process.")

    # Row 1: Key to Search
    Label(flat_tab, text="?", bg="blue", fg="white", font=("Arial", 8, "bold")).grid(row=1, column=0, padx=2, pady=5,
                                                                                     sticky="e")
    Label(flat_tab, text="Key to Search:").grid(row=1, column=1, padx=10, pady=5, sticky="w")
    key_entry = Entry(flat_tab, width=40)
    key_entry.grid(row=1, column=2, padx=10, pady=5)
    CreateToolTip(flat_tab.grid_slaves(row=1, column=0)[0], "Enter the key (e.g., VersionsNr) you wish to search for.")

    # Row 2: New Value
    Label(flat_tab, text="?", bg="blue", fg="white", font=("Arial", 8, "bold")).grid(row=2, column=0, padx=2, pady=5,
                                                                                     sticky="e")
    Label(flat_tab, text="New Value:").grid(row=2, column=1, padx=10, pady=5, sticky="w")
    value_entry = Entry(flat_tab, width=40)
    value_entry.grid(row=2, column=2, padx=10, pady=5)
    CreateToolTip(flat_tab.grid_slaves(row=2, column=0)[0], "Enter the new value (e.g., to 1.0.6) to replace the current value (e.g., from 1.0.7).")

    # Row 3: Remove Key Checkbox
    Label(flat_tab, text="?", bg="blue", fg="white", font=("Arial", 8, "bold")).grid(row=3, column=0, padx=2, pady=5,
                                                                                     sticky="e")
    Label(flat_tab, text="Remove Key:").grid(row=3, column=1, padx=10, pady=5, sticky="w")
    remove_key_var = IntVar()
    Checkbutton(flat_tab, variable=remove_key_var).grid(row=3, column=2, padx=10, pady=5, sticky="w")
    CreateToolTip(flat_tab.grid_slaves(row=3, column=0)[0], "Check this if you want to remove the key and its value entirely.")

    # Row 4: Log widget
    log_widget = Text(flat_tab, height=15, width=70)
    log_widget.grid(row=4, column=0, columnspan=4, padx=10, pady=10)
    scrollbar = Scrollbar(flat_tab, orient=VERTICAL, command=log_widget.yview)
    scrollbar.grid(row=4, column=4, sticky="ns")
    log_widget.config(yscrollcommand=scrollbar.set)
    log_widget.tag_config("error", foreground="red")
    log_widget.tag_config("warning", foreground="orange")
    log_widget.tag_config("success", foreground="green")
    log_widget.tag_config("info", foreground="blue")
    log_widget.insert(END, f"🛠 Version: {VERSION}\n", "info")

    # Row 5: Progress bar and row 6: Percentage label
    progress_bar = Progressbar(flat_tab, orient="horizontal", mode="determinate", length=400)
    progress_bar.grid(row=5, column=0, columnspan=4, padx=10, pady=5)
    progress_bar.grid_remove()
    percent_label = Label(flat_tab, text="")
    percent_label.grid(row=6, column=0, columnspan=4, pady=5)
    percent_label.grid_remove()

    flat_tab.grid_columnconfigure(0, weight=1, uniform="col")
    flat_tab.grid_columnconfigure(1, weight=1, uniform="col")
    flat_tab.grid_columnconfigure(2, weight=1, uniform="col")
    flat_tab.grid_columnconfigure(3, weight=1, uniform="col")

    Button(flat_tab, text="Start Processing", command=lambda: start_processing(
        directory_entry, key_entry, value_entry, remove_key_var, log_widget, progress_bar, percent_label,
        backup_entry_settings.get()
    )).grid(row=7, column=1, columnspan=2, pady=10)

    Button(flat_tab, text="Check for Updates", command=check_for_updates).grid(row=8, column=1, columnspan=2, pady=5)

    github_label = Label(flat_tab, text="View on GitHub", fg="blue", cursor="hand2")
    github_label.grid(row=9, column=1, columnspan=2, pady=10)
    github_label.bind("<Button-1>", lambda e: open_github_link())

    # ----- Cell value fixer Tab (Dynamic Replacement) -----
    cell_tab = Frame(notebook)
    notebook.add(cell_tab, text="Cell value fixer")

    Label(cell_tab, text="?", bg="blue", fg="white", font=("Arial", 8, "bold")).grid(row=0, column=0, padx=2, pady=5,
                                                                                     sticky="e")
    Label(cell_tab, text="Select Processing Directory:").grid(row=0, column=1, padx=10, pady=5, sticky="w")
    directory_entry_value = Entry(cell_tab, width=40)
    directory_entry_value.grid(row=0, column=2, padx=10, pady=5)
    Button(cell_tab, text="Browse", command=lambda: browse_directory(directory_entry_value)).grid(row=0, column=3,
                                                                                                  padx=10, pady=5)
    CreateToolTip(cell_tab.grid_slaves(row=0, column=0)[0],
                  "Choose the folder containing Excel files for value replacement.")

    Label(cell_tab, text="?", bg="blue", fg="white", font=("Arial", 8, "bold")).grid(row=1, column=0, padx=2, pady=5,
                                                                                     sticky="e")
    Label(cell_tab, text="Current Value:").grid(row=1, column=1, padx=10, pady=5, sticky="w")
    current_value_entry = Entry(cell_tab, width=40)
    current_value_entry.grid(row=1, column=2, padx=10, pady=5)
    CreateToolTip(cell_tab.grid_slaves(row=1, column=0)[0], "Enter the exact text you want to replace.")

    Label(cell_tab, text="?", bg="blue", fg="white", font=("Arial", 8, "bold")).grid(row=2, column=0, padx=2, pady=5,
                                                                                     sticky="e")
    Label(cell_tab, text="New Value:").grid(row=2, column=1, padx=10, pady=5, sticky="w")
    new_value_entry_value = Entry(cell_tab, width=40)
    new_value_entry_value.grid(row=2, column=2, padx=10, pady=5)
    CreateToolTip(cell_tab.grid_slaves(row=2, column=0)[0], "Enter the new text to insert.")

    log_widget_value = Text(cell_tab, height=10, width=70)
    log_widget_value.grid(row=3, column=0, columnspan=4, padx=10, pady=10)
    scrollbar_value = Scrollbar(cell_tab, orient=VERTICAL, command=log_widget_value.yview)
    scrollbar_value.grid(row=3, column=4, sticky="ns")
    log_widget_value.config(yscrollcommand=scrollbar_value.set)
    log_widget_value.tag_config("error", foreground="red")
    log_widget_value.tag_config("warning", foreground="orange")
    log_widget_value.tag_config("success", foreground="green")
    log_widget_value.tag_config("info", foreground="blue")

    progress_bar_value = Progressbar(cell_tab, orient="horizontal", mode="determinate", length=400)
    progress_bar_value.grid(row=4, column=0, columnspan=4, padx=10, pady=5)
    progress_bar_value.grid_remove()
    percent_label_value = Label(cell_tab, text="")
    percent_label_value.grid(row=5, column=0, columnspan=4, pady=5)
    percent_label_value.grid_remove()

    cell_tab.grid_columnconfigure(0, weight=1, uniform="col")
    cell_tab.grid_columnconfigure(1, weight=1, uniform="col")
    cell_tab.grid_columnconfigure(2, weight=1, uniform="col")
    cell_tab.grid_columnconfigure(3, weight=1, uniform="col")

    Button(cell_tab, text="Replace Value", command=lambda: threading.Thread(target=start_value_replacement, args=(
        directory_entry_value.get(),
        current_value_entry.get(),
        new_value_entry_value.get(),
        log_widget_value,
        progress_bar_value,
        percent_label_value,
        backup_entry_settings.get()
    )).start()).grid(row=6, column=1, columnspan=2, pady=10)

    # ----- Settings Tab (Backup, Check Updates, GitHub) -----
    settings_tab = Frame(notebook)
    notebook.add(settings_tab, text="Settings")

    Label(settings_tab, text="?", bg="blue", fg="white", font=("Arial", 8, "bold")).grid(row=0, column=0, padx=2,
                                                                                         pady=5, sticky="e")
    Label(settings_tab, text="Select Backup Directory:").grid(row=0, column=1, padx=10, pady=5, sticky="w")
    global backup_entry_settings
    backup_entry_settings = Entry(settings_tab, width=40)
    backup_entry_settings.grid(row=0, column=2, padx=10, pady=5)
    Button(settings_tab, text="Browse", command=lambda: browse_directory(backup_entry_settings)).grid(row=0, column=3,
                                                                                                      padx=10, pady=5)
    CreateToolTip(settings_tab.grid_slaves(row=0, column=0)[0],
                  "Choose the folder where backup copies of your Excel files will be stored.")

    settings_tab.grid_columnconfigure(0, weight=1, uniform="col")
    settings_tab.grid_columnconfigure(1, weight=1, uniform="col")
    settings_tab.grid_columnconfigure(2, weight=1, uniform="col")
    settings_tab.grid_columnconfigure(3, weight=1, uniform="col")

    Button(settings_tab, text="Check for Updates", command=check_for_updates).grid(row=1, column=1, columnspan=2,
                                                                                   pady=10)
    github_label_settings = Label(settings_tab, text="View on GitHub", fg="blue", cursor="hand2")
    github_label_settings.grid(row=2, column=1, columnspan=2, pady=10)
    github_label_settings.bind("<Button-1>", lambda e: open_github_link())

    root.mainloop()


if __name__ == "__main__":
    launch_gui()
