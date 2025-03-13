# Excel_word_changer
GUI app that is replacing selected word in one or many xlsx (excel related) files. 


### How to use the GUI:
1. Run the Script: Run the script using Python, and a GUI window will open.
2. Directory Selection: Click "Browse" to select the directory containing the Excel files.
3. Key and Value: Enter the key to search for (e.g., LocId) and the new value (e.g., R00).
4. Start Processing: Click "Start Processing" to begin.
5. Progress Logs: The log area will display updates and results of the processing.

## Important! ## 
For now the application works only with files that have cells with pipe separated keys and values. For example: |a=1|b=2|c=3|

### How to export it as an executable:
Follow the steps from the earlier explanation to convert this script into a standalone executable:
1. Install pyinstaller:
    `pip install pyinstaller`
2. Package the script: 
    `pyinstaller --onefile --windowed --hidden-import=openpyxl --icon=excellKilla.ico xlsx_GUI.py`
--windowed ensures that no console window appears when the app runs.
3. Share the .exe file located in the dist folder.


### How it can help us in real life?
The application will:
* Provide an intuitive interface for selecting directories and specifying values.
* Process Excel files in the selected folder and its subfolders.
* Display progress and errors in the log area.
