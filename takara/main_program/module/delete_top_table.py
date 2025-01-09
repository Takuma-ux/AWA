import os
import json
import re
import argparse
import win32com.client as win32

def remove_before_first_heading1(document, output_file_path):
    # ドキュメントの最初の位置を取得
    range_start = document.Content.Start
    range_end = None

    # 最初の見出し1スタイルの段落を検索
    for paragraph in document.Paragraphs:
        if paragraph.Range.Style is not None and paragraph.Range.Style.NameLocal == "見出し 1":
            range_end = paragraph.Range.Start  # 見出し1の開始位置までを範囲として設定
            break

    # 範囲のデバッグ情報を表示
    print(f"range_start: {range_start}, range_end: {range_end}")

    if range_end is not None:
        # 指定範囲を削除
        document.Range(Start=range_start, End=range_end).Delete()
        print("指定された範囲が削除されました。")
    else:
        print("見出し1が見つかりませんでした。")

    # 変更を保存
    document.SaveAs(output_file_path)
    print(f"変更が保存されました: {output_file_path}")

# 使用例:
# コマンドライン引数をパースするための設定
parser = argparse.ArgumentParser(description='Process some files.')
parser.add_argument('--config', required=True, help='Path to the config JSON file')
args = parser.parse_args()

# JSONファイルのパスを取得
json_file_path = args.config

# JSONファイルを開いて読み込む
with open(json_file_path, 'r', encoding='utf-8-sig') as file:
    data = json.load(file)

base_dir = os.path.dirname(json_file_path)

# 元のdocxファイルのパスを取得
docx_raw_file_path = os.path.abspath(os.path.join(base_dir, data["docx_raw_file_path"]))

# JSONファイル名から数字を抽出
match = re.search(r'\d+', os.path.basename(json_file_path))
if match:
    number = match.group()
else:
    number = 'default'  # 数字が見つからない場合のデフォルト値

output_dir = os.path.join(base_dir, "output", f"{number}")
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

docx_raw_file_name = os.path.basename(docx_raw_file_path)
docx_file_name_modified = docx_raw_file_name.replace('.docx', '_without_toc.docx')
docx_file_path_2 = os.path.abspath(os.path.join(output_dir, docx_file_name_modified))

# 出力ファイルのパスを設定
output_file_path = docx_file_path_2.replace('.docx', '_final.docx')

# ファイルを開く
word = win32.Dispatch("Word.Application")
word.Visible = False  # 非表示で処理を実行

try:
    # Wordドキュメントを開く
    document = word.Documents.Open(docx_file_path_2, ReadOnly=False)

    # 指定範囲を削除する関数を実行
    remove_before_first_heading1(document, output_file_path)

    # ドキュメントを閉じて保存
    document.Close(SaveChanges=True)

finally:
    word.Quit()
