@echo off
echo PyInstallerでexeファイルをビルドしています...

REM 仮想環境をアクティベートする（存在する場合）
if exist venv\Scripts\activate.bat (
    echo 仮想環境をアクティベートしています...
    call venv\Scripts\activate.bat
)

REM 必要なパッケージをインストール
echo 必要なパッケージをインストールしています...
pip install -r requirements.txt
pip install pyinstaller

REM 古いビルドファイルを削除
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

REM PyInstallerでビルド
echo exeファイルをビルドしています...
pyinstaller build.spec

if exist dist\excel_image_converter.exe (
    echo ビルドが完了しました！
    echo exeファイルは dist\excel_image_converter.exe にあります。
    echo.
    echo 使用方法:
    echo   excel_image_converter.exe -i 入力フォルダ -o 出力フォルダ
    echo   例: excel_image_converter.exe -i excels -o images
) else (
    echo ビルドに失敗しました。
)

pause