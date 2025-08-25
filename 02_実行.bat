@echo off
echo Excelの画像変換処理を開始します...
echo.

call venv\Scripts\activate
python convert_excel_to_images.py

echo.
echo 全ての処理が完了しました。
pause