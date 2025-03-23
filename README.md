# xlsx fixer

xlsx fixer is a Python-based GUI application designed to process Excel (.xlsx) files. It allows you to search for a
specified key in the Excel files and either replace its value or remove the key entirely. Before making any
modifications, the application creates a backup of your original files—preserving the folder structure—to ensure that
your data is safe.

### Features

#### Backup Functionality:
* Automatically copies all Excel files from your processing directory into a backup folder.
* If no backup directory is specified, a folder named "Backup" will be created inside the processing directory.
* If you mistakenly choose the processing directory as the backup location, the app automatically uses a subfolder named "Backup" instead.

#### Search and Replace / Removal:
* Search for a specified key in Excel files.
* Replace the key’s value with a new value or completely remove the key-value pair.

#### Progress Tracking:
* Displays a progress bar while processing large numbers of files.
* Logs details of processed files and any errors encountered.

#### User-Friendly GUI:
* Built with Tkinter, includes buttons for browsing directories, starting the process, checking for updates, and linking to GitHub.
* A splash screen is shown at startup for a modern, professional look.

#### Version Checking:
* the application checks for the latest version on GitHub and notifies you if an update is available.



### Installation

Follow the steps from the earlier explanation to convert this script into a standalone executable:

**1\. Clone the repository:**

    git clone https://github.com/Maksymilianx/Excel_word_changer.git
    cd Excel_word_changer
   
**2\. Set up your environment and install dependencies:** 

It is recommended to use a virtual environment.
 
    python -m venv venv
    source venv/bin/activate      # On Windows: venv\Scripts\activate
    pip install -r requirements.txt

   Note: The `requirements.txt` file should include openpyxl, requests, and any other required libraries.

**3\. Run the application:**

    python xlsx_GUI.py

### Packaging as a Standalone Executable
To package the application into a standalone executable (no Python installation required for users):

1\. Install PyInstaller (if not already installed):

    pip install pyinstaller

2\. Create the executable:

Replace your_icon.ico with the path to your actual icon file:

    pyinstaller --onefile --windowed --hidden-import=openpyxl --icon=your_icon.ico xlsx_GUI.py


3\. Locate the executable:

Your executable will be created inside the dist folder.
You can now distribute this executable to users who don't have Python installed.

### Usage

Follow these steps to use the xlxs fixer application:

1. Select Backup Directory:
   * Either select a backup folder or leave this field empty.
   * If left empty, a folder named "Backup" will be created inside your processing directory.
   * If you select the processing directory as backup, a subfolder "Backup" is automatically created.
2. Select Processing Directory:
   * Select the folder containing the Excel files you wish to modify.
3. Enter Key and New Value (or select "Remove Key"):
   * Enter the key you want to find.
   * Provide a new value for replacing, or check "Remove Key" to delete the key entirely.
4. Start Processing:
   * Click "Start Processing".
   * The app first backs up your original files and then processes them.
   * Monitor the progress via the progress bar and logs displayed in the GUI.
5. Additional Options:
   * Click "Check for Updates" to see if a newer version is available.
   * Click the "View on GitHub" link to open the repository in your browser.

### Important!  

For now the application works only with files that have cells with pipe separated keys and values. For example:
|a=1|b=2|c=3|

### Contributing
Contributions are welcome! To contribute to the project:
* Fork the repository.
* Create a new branch (git checkout -b feature/your-feature-name).
* Commit your changes (git commit -m 'Add some feature').
* Push to your branch (git push origin feature/your-feature-name).
* Open a new Pull Request on GitHub. 
 
Report issues or feature requests by creating an [issue](https://github.com/Maksymilianx/Excel_word_changer/issues).
