# Excelシート画像化ツール

このツールは、`excels`フォルダに入れた複数のExcelファイルの全シートを、一括で画像ファイル（PNG形式）に変換します。

## 動作環境の準備（初回のみ）

このツールを初めて使用する前に、お使いのOSに合わせて以下の準備を一度だけ行ってください。

---
### Windowsの場合 🖥️

1.  **Microsoft Excel**
    * PCにインストールされている必要があります。

2.  **Python**
    * [Python公式サイト](https://www.python.org/downloads/windows/)からインストーラーをダウンロードします。
    * **重要:** インストーラー実行時、必ず「**Add Python to PATH**」のチェックボックスにチェックを入れてください。
    

3.  **Poppler (PDF変換ツール)**
    * [こちらのリンク](https://github.com/oschwartz10612/poppler-windows/releases/)から最新版のZIPファイルをダウンロードし、解凍します。
    * 解凍してできたフォルダの中にある`bin`フォルダ（例: `C:\poppler\bin`）のパスを、Windowsの**環境変数 `Path`** に追加してください。

4.  **Pythonライブラリ**
    * コマンドプロンプトを開き、以下のコマンドを実行して必要なライブラリをインストールします。
    ```sh
    pip install pdf2image pywin32
    ```

---
### macOSの場合 

1.  **Microsoft Excel**
    * PCにインストールされている必要があります。

2.  **Homebrew**
    * まだインストールしていない場合は、ターミナルを開き、以下のコマンドを一度だけ実行します。
    ```sh
    /bin/bash -c "$(curl -fsSL [https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh](https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh))"
    ```

3.  **各種ツール**
    * ターミナルで以下のコマンドを実行し、必要なツールをインストールします。
    ```sh
    brew install poppler
    ```

4.  **Pythonライブラリ**
    * ターミナルで以下のコマンドを実行して必要なライブラリをインストールします。
    ```sh
    pip install pdf2image
    ```