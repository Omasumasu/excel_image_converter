#!/bin/bash
# macOS用の実行スクリプト

echo "Excelの画像変換処理を開始します..."
echo ""

source venv/bin/activate
python3 convert_excel_to_images.py

echo ""
echo "全ての処理が完了しました。"
read -p "何かキーを押して終了します..."