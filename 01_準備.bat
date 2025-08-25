@echo off
echo ツールに必要なフォルダとライブラリを準備します...
echo.

echo フォルダを作成中...
if not exist "excels" mkdir excels
if not exist "images" mkdir images

echo ライブラリをインストール中...
python -m venv venv
call venv\Scripts\activate
pip install pdf2image pywin32

echo.
echo 準備が完了しました。
pause