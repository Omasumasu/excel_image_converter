@echo off
echo ツールに必要なライブラリを準備します...
echo.

python -m venv venv
call venv\Scripts\activate
pip install pdf2image pywin32

echo.
echo 準備が完了しました。
pause