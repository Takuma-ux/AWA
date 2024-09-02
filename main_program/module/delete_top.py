import os
import win32com.client as win32
import json
import argparse

# Word の定数を定義
wdBorderTop = 1
wdBorderBottom = 3
wdBorderLeft = 4
wdBorderRight = 2
wdWithInTable = 12

def remove_table_of_contents(file_path):
    # Word アプリケーションの起動
    word = win32.Dispatch("Word.Application")
    word.Visible = False  # Wordを非表示で実行

    # ドキュメントを編集可能モードで開く
    document = word.Documents.Open(file_path)

    try:
        # 目次をすべて削除
        for toc in document.TablesOfContents:
            toc.Delete()

        # 変更を保存
        # outputディレクトリに保存するように変更
        output_dir = os.path.join(os.path.dirname(file_path), '..', 'output')
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        output_file_name = os.path.basename(file_path).replace('.docx', '_without_toc.docx')
        output_file_path = os.path.abspath(os.path.join(output_dir, output_file_name))
        
        document.SaveAs(output_file_path)
        print(f"変更が保存されました: {output_file_path}")
    except Exception as e:
        print(f"エラーが発生しました: {e}")
    finally:
        document.Close(False)
        word.Quit()

# 使用例
# コマンドライン引数をパースするための設定
parser = argparse.ArgumentParser(description='Process some files.')
parser.add_argument('--config', required=True, help='Path to the config JSON file')
args = parser.parse_args()

# JSONファイルのパスを取得
json_file_path = args.config
# JSONファイルで指定されている元のパスを取得
# JSONファイルを開いて読み込む
with open(json_file_path, 'r', encoding='utf-8-sig') as file:
    data = json.load(file)

base_dir = os.path.dirname(json_file_path) 
# 元のdocxファイルのパスを取得
docx_raw_file_path = os.path.abspath(os.path.join(base_dir, data["docx_raw_file_path"]))

remove_table_of_contents(docx_raw_file_path)
