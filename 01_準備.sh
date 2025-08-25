#!/bin/bash
# macOS用の準備スクリプト

echo "ツールに必要なフォルダ、ツール、ライブラリを準備します..."
echo ""

echo "フォルダを作成します..."
mkdir -p excels
mkdir -p images

# Homebrewのインストールチェック
if ! command -v brew &> /dev/null
then
    echo "Homebrewが見つかりません。インストールを開始します..."
    /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
fi

echo "Homebrewを使って必要なツールをインストールします..."
brew install poppler

echo "Pythonの仮想環境を作成し、ライブラリをインストールします..."
python3 -m venv venv
source venv/bin/activate
pip install pdf2image

echo ""
echo "準備が完了しました。次回以降は '02_実行.sh' を使ってください。"
read -p "何かキーを押して終了します..."