from tkinter import Tk, Label, Entry, Button, Text, Scrollbar, VERTICAL, END, IntVar, Checkbutton
from tkinter.ttk import Progressbar
from functions import (
    check_for_updates,
    browse_directory,
    open_github_link,
    start_processing,
    VERSION
)

def launch_gui():
    root = Tk()
    root.title(f"xlxs fixer - v{VERSION}")

    # Backup Directory selection (Row 0)
    Label(root, text="Select Backup Directory:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
    backup_entry = Entry(root, width=50)
    backup_entry.grid(row=0, column=1, padx=10, pady=5)
    Button(root, text="Browse", command=lambda: browse_directory(backup_entry)).grid(row=0, column=2, padx=10, pady=5)

    # Processing Directory selection (Row 1)
    Label(root, text="Select Processing Directory:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
    directory_entry = Entry(root, width=50)
    directory_entry.grid(row=1, column=1, padx=10, pady=5)
    Button(root, text="Browse", command=lambda: browse_directory(directory_entry)).grid(row=1, column=2, padx=10, pady=5)

    # Key input (Row 2)
    Label(root, text="Key to Search:").grid(row=2, column=0, padx=10, pady=5, sticky="w")
    key_entry = Entry(root, width=50)
    key_entry.grid(row=2, column=1, padx=10, pady=5)

    # New value input (Row 3)
    Label(root, text="New Value:").grid(row=3, column=0, padx=10, pady=5, sticky="w")
    value_entry = Entry(root, width=50)
    value_entry.grid(row=3, column=1, padx=10, pady=5)

    # Remove Key Checkbox (Row 4)
    remove_key_var = IntVar()
    Checkbutton(root, text="Remove Key", variable=remove_key_var).grid(row=4, column=1, pady=5)

    # Log display (Row 5)
    log_widget = Text(root, height=15, width=70)
    log_widget.grid(row=5, column=0, columnspan=3, padx=10, pady=10)
    scrollbar = Scrollbar(root, orient=VERTICAL, command=log_widget.yview)
    scrollbar.grid(row=5, column=3, sticky="ns")
    log_widget.config(yscrollcommand=scrollbar.set)
    log_widget.tag_config("error", foreground="red")
    log_widget.tag_config("warning", foreground="orange")
    log_widget.tag_config("success", foreground="green")
    log_widget.tag_config("info", foreground="blue")
    log_widget.insert(END, f"🛠 Version: {VERSION}\n", "info")

    # Progress Bar (Row 6)
    progress_bar = Progressbar(root, orient="horizontal", mode="determinate", length=400)
    progress_bar.grid(row=6, column=0, columnspan=3, padx=10, pady=10)

    # Start Processing Button (Row 7)
    Button(root, text="Start Processing", command=lambda: start_processing(
        directory_entry, backup_entry, key_entry, value_entry, remove_key_var, log_widget, progress_bar)
    ).grid(row=7, column=1, pady=10)

    # Check for Updates button (Row 8)
    Button(root, text="Check for Updates", command=check_for_updates).grid(row=8, column=1, pady=5)

    # GitHub link (Row 9)
    github_label = Label(root, text="View on GitHub", fg="blue", cursor="hand2")
    github_label.grid(row=9, column=1, pady=10)
    github_label.bind("<Button-1>", lambda e: open_github_link())

    root.mainloop()

if __name__ == "__main__":
    launch_gui()
