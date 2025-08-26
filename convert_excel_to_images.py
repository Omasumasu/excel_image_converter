import os
import subprocess
import tempfile
import platform
import argparse
from pdf2image import convert_from_path

# --- Windows専用のExcel操作関数 ---
def convert_excel_to_pdf_windows(excel_path, output_folder):
    import win32com.client
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        workbook = excel.Workbooks.Open(excel_path)
        base_name = os.path.splitext(os.path.basename(excel_path))[0]
        for sheet in workbook.Worksheets:
            sheet.Activate()
            sheet.PageSetup.Zoom = False
            sheet.PageSetup.FitToPagesWide = 1
            sheet.PageSetup.FitToPagesTall = 1
            pdf_path = os.path.join(output_folder, f"{base_name}---{sheet.Name}.pdf")
            workbook.ActiveSheet.ExportAsFixedFormat(0, pdf_path)
        workbook.Close(SaveChanges=False)
        print(f"  [成功] '{os.path.basename(excel_path)}' の全シートをPDFに変換しました。")
        return True
    except Exception as e:
        print(f"  [エラー] WindowsでのExcel操作に失敗しました: {e}")
        return False
    finally:
        if excel:
            excel.Quit()

# --- macOS専用のExcel操作関数 ---
def convert_excel_to_pdf_mac(excel_path, output_folder):
    base_name = os.path.splitext(os.path.basename(excel_path))[0]
    apple_script = f'''
    tell application "Microsoft Excel"
        set screen updating to false
        set excelFile to POSIX file "{excel_path}"
        open excelFile
        set workbook_sheets to every sheet of active workbook
        repeat with current_sheet in workbook_sheets
            activate object current_sheet
            try
                set pageSetupObject to page setup of active sheet
                set fit to pages wide of pageSetupObject to 1
                set fit to pages tall of pageSetupObject to 1
            on error
                log "ページ設定はスキップされました。"
            end try
            set current_sheet_name to name of current_sheet
            set output_path to "{output_folder}/" & "{base_name}" & "---" & current_sheet_name & ".pdf"
            save active sheet in (POSIX file output_path) as PDF file format
        end repeat
        close active workbook without saving
        set screen updating to true
    end tell
    '''
    try:
        subprocess.run(['osascript', '-e', apple_script], check=True, capture_output=True, text=True, timeout=180)
        print(f"  [成功] '{os.path.basename(excel_path)}' の全シートをPDFに変換しました。")
        return True
    except Exception as e:
        print(f"  [エラー] macOSでのExcel操作に失敗しました: {e}")
        return False

# --- メイン処理（OS共通） ---
def convert_all_excels_to_images(input_dir="excels", output_dir="images"):
    if not os.path.isdir(output_dir):
        os.makedirs(output_dir)
    print(f"'{input_dir}' ディレクトリ内のExcelファイルを画像に変換します...")
    current_os = platform.system()
    print(f"現在のOS: {current_os}")

    poppler_binary_path = None
    if current_os == "Windows":
        # The batch script renames the folder to "poppler"
        poppler_path = os.path.abspath("poppler")
        if os.path.isdir(poppler_path):
            poppler_binary_path = os.path.join(poppler_path, 'Library/bin')
            print(f"Popplerのパスを検出しました: {poppler_binary_path}")

    with tempfile.TemporaryDirectory() as temp_dir:
        for filename in os.listdir(input_dir):
            if filename.lower().endswith('.xlsx') and not filename.startswith('~'):
                excel_filepath = os.path.abspath(os.path.join(input_dir, filename))
                print(f"\n--- ファイル '{filename}' を処理中 ---")
                pdf_converted = False
                if current_os == "Windows":
                    pdf_converted = convert_excel_to_pdf_windows(excel_filepath, temp_dir)
                elif current_os == "Darwin":
                    pdf_converted = convert_excel_to_pdf_mac(excel_filepath, temp_dir)
                else:
                    print(f"  [エラー] サポートされていないOSです: {current_os}")
                    continue
                if not pdf_converted:
                    continue
                for pdf_file in os.listdir(temp_dir):
                    if pdf_file.startswith(os.path.splitext(filename)[0]):
                        pdf_filepath = os.path.join(temp_dir, pdf_file)
                        try:
                            images = convert_from_path(pdf_filepath, poppler_path=poppler_binary_path)
                            for i, img in enumerate(images):
                                base_pdf_name = os.path.splitext(pdf_file)[0]
                                output_image_name = f"{base_pdf_name}_page_{i+1}.png"
                                output_image_path = os.path.join(output_dir, output_image_name)
                                img.save(output_image_path, "PNG")
                            print(f"  [情報] '{pdf_file}' から {len(images)} 枚の画像を保存しました。")
                        except Exception as e:
                            print(f"  [エラー] '{pdf_file}' から画像への変換に失敗しました: {e}")
    print("\n全ての変換処理が完了しました。")

def main():
    parser = argparse.ArgumentParser(description='Excel files to images converter')
    parser.add_argument('-i', '--input', default='excels', 
                       help='Input directory containing Excel files (default: excels)')
    parser.add_argument('-o', '--output', default='images',
                       help='Output directory for generated images (default: images)')
    
    args = parser.parse_args()
    
    print(f"入力フォルダ: {args.input}")
    print(f"出力フォルダ: {args.output}")
    
    if not os.path.exists(args.input):
        print(f"エラー: 入力フォルダ '{args.input}' が見つかりません。")
        return 1
    
    convert_all_excels_to_images(args.input, args.output)
    return 0

if __name__ == "__main__":
    import sys
    sys.exit(main())