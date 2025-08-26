#!/bin/bash
echo "PyInstallerでexeファイルをビルドしています..."

# 仮想環境をアクティベートする（存在する場合）
if [ -f "venv/bin/activate" ]; then
    echo "仮想環境をアクティベートしています..."
    source venv/bin/activate
fi

# 必要なパッケージをインストール
echo "必要なパッケージをインストールしています..."
pip install -r requirements.txt
pip install pyinstaller

# 古いビルドファイルを削除
rm -rf build dist

# PyInstallerでビルド
echo "exeファイルをビルドしています..."
pyinstaller build.spec

if [ -f "dist/excel_image_converter" ]; then
    echo "ビルドが完了しました！"
    echo "実行ファイルは dist/excel_image_converter にあります。"
    echo ""
    echo "使用方法:"
    echo "  ./dist/excel_image_converter -i 入力フォルダ -o 出力フォルダ"
    echo "  例: ./dist/excel_image_converter -i excels -o images"
else
    echo "ビルドに失敗しました。"
fi