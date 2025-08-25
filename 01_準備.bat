@echo off
echo =======================================================
echo  Excel画像変換ツール 準備プログラム
echo =======================================================
echo.
echo ツールに必要なフォルダとライブラリを準備します...
echo.

echo --- 1. フォルダを作成中...
if not exist "excels" mkdir excels
if not exist "images" mkdir images
echo    -> 完了
echo.

echo --- 2. PDF変換ツール(Poppler)をダウンロード中...
echo    (これには少し時間がかかる場合があります)
curl -L "https://github.com/oschwartz10612/poppler-windows/releases/download/v25.07.0-0/Release-25.07.0-0.zip" -o "poppler.zip"
echo    -> 完了
echo.

echo --- 3. Popplerを展開中...
tar -xf poppler.zip
rename Release-25.07.0-0 poppler
echo    -> 完了
echo.

echo --- 4. PopplerのZIPファイルを削除中...
del poppler.zip
echo    -> 完了
echo.

echo --- 5. Pythonライブラリをインストール中...
python -m venv venv
call venv\Scripts\activate
pip install pdf2image pywin32
echo    -> 完了
echo.
echo =======================================================
echo  全ての準備が完了しました！
echo =======================================================
echo.
pause