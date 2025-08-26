@echo off
rem --- Setup Script for Excel Image Converter ---

echo =======================================================
echo  Starting setup for the Excel Image Converter...
echo =======================================================
echo.

echo --- 1. Creating folders...
if not exist "excels" mkdir excels
if not exist "images" mkdir images
echo    Done.
echo.

echo --- 2. Downloading PDF converter (Poppler)...
echo    (This may take a moment)
curl -L "https://github.com/oschwartz10612/poppler-windows/releases/download/v25.07.0-0/Release-25.07.0-0.zip" -o "poppler.zip"
echo    Done.
echo.

echo --- 3. Unzipping Poppler...
tar -xf poppler.zip
echo    Done.
echo.

echo --- 4. Renaming Poppler folder...
if exist "poppler-25.07.0" (
    rename "poppler-25.07.0" "poppler"
    echo    Done.
) else (
    echo    Folder not found, skipping rename.
)
echo.

echo --- 5. Deleting ZIP file...
del poppler.zip
echo    Done.
echo.

echo --- 6. Installing Python libraries...
python -m venv venv
call venv\Scripts\activate
pip install pdf2image pywin32
echo    Done.
echo.
echo =======================================================
echo  Setup has been completed successfully!
echo =======================================================
echo.
pause