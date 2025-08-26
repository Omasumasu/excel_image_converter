@echo off
rem --- Execution Script for Excel Image Converter ---

echo Starting the Excel to Image conversion process...
echo.

call venv\Scripts\activate
python convert_excel_to_images.py

echo.
echo All processes have been completed.
pause