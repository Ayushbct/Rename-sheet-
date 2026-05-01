# Excel Sheet Processor

A tool to automatically rename Excel sheets and clean data in columns Z (engineno) and AA (chasisno).

## Features

✅ User-friendly GUI (no command line needed)  
✅ Browse and select folder containing Excel files  
✅ Process multiple Excel files at once  
✅ Clean columns Z & AA (remove spaces and hyphens)  
✅ Rename sheets to "Sheet1"  
✅ Real-time processing log  
✅ Works with .xlsx and .xls files  

## How to Create the .exe File

### Prerequisites
- Windows OS
- Python 3.7+ installed
- Command Prompt or PowerShell

### Step-by-Step Instructions

1. **Install required packages:**
   ```cmd
   pip install -r requirements.txt
   ```

2. **Build the .exe file:**
   - **Option A (Automatic):** Double-click `build.bat`
   - **Option B (Manual):** Run in Command Prompt:
     ```cmd
     pyinstaller --onefile --windowed --name="Excel_Processor" excel_processor_gui.py
     ```

3. **Find the .exe file:**
   - After building, the .exe will be in the `dist` folder
   - File name: `Excel_Processor.exe`

4. **Use it:**
   - Copy `Excel_Processor.exe` to any location
   - Double-click to run
   - Click "Browse Folder" to select a folder with Excel files
   - Click "Process Excel Files" to start

## Usage

1. Run `Excel_Processor.exe`
2. Click "Browse Folder" button
3. Select the folder containing your Excel files
4. Click "Process Excel Files"
5. View the processing log in real-time
6. When complete, your files will be updated

## What It Does

For each Excel file:
- ✓ Renames the sheet to "Sheet1" (if it's not already)
- ✓ Cleans column Z (engineno) - removes spaces and hyphens
- ✓ Cleans column AA (chasisno) - removes spaces and hyphens
- ✓ Saves the modified file

## File Structure

```
Rename-sheet-/
├── excel_processor_gui.py        (GUI Application)
├── main.py                       (Command-line version)
├── build.bat                     (Build script for .exe)
├── requirements.txt              (Dependencies)
└── dist/
    └── Excel_Processor.exe       (Final .exe file)
```