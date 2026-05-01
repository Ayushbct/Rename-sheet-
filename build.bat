@echo off
REM Build Script for Excel Processor GUI

echo Installing required packages...
pip install -r requirements.txt

echo.
echo Building .exe file...
pyinstaller --onefile --windowed --icon=icon.ico --name="Excel_Processor" excel_processor_gui.py

echo.
echo Build complete! 
echo The .exe file is located in the 'dist' folder
echo File: dist\Excel_Processor.exe
pause
