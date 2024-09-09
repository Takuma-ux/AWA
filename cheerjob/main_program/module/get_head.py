import win32com.client
import os
import json
import argparse
import re

def is_heading1(word_range):
    if word_range is None:
        return False
    if word_range.Style is not None:
        return word_range.Style.NameLocal == "見出し 1"
    return False

def is_heading2(word_range):
    if word_range is None:
        return False
    if word_range.Style is not None:
        return word_range.Style.NameLocal == "見出し 2"
    return False

# Wordアプリケーションを起動
word = win32com.client.Dispatch("Word.Application")
word.Visible = False  # Wordアプリケーションを非表示にする


# 出力ファイルのパスを設定
# コマンドライン引数をパースするための設定
parser = argparse.ArgumentParser(description='Process some files.')
parser.add_argument('--config', required=True, help='Path to the config JSON file')
args = parser.parse_args()

# JSONファイルのパスを取得
json_file_path = args.config
with open(json_file_path, 'r', encoding='utf-8-sig') as file:
    data = json.load(file)

base_dir = os.path.dirname(json_file_path) 
# 元のdocxファイルのパスを取得
docx_raw_file_path = os.path.abspath(os.path.join(base_dir, data["docx_raw_file_path"]))
heading1_file_path = os.path.abspath(os.path.join(base_dir, data["heading1_file_path"]))
heading2_file_path = os.path.abspath(os.path.join(base_dir, data["heading2_file_path"]))
# docx_file_path_2 を生成
docx_raw_file_name = os.path.basename(docx_raw_file_path)
# JSONファイル名から数字を抽出
match = re.search(r'\d+', os.path.basename(json_file_path))
if match:
    number = match.group()
else:
    number = 'default'  # 数字が見つからない場合のデフォルト値

output_dir = os.path.join(base_dir, "output", f"{number}")
if not os.path.exists(output_dir):
    os.makedirs(output_dir)
docx_file_name_modified = docx_raw_file_name.replace('.docx', '_without_toc.docx')
docx_file_path_2 = os.path.join(output_dir, docx_file_name_modified)
doc = word.Documents.Open(docx_raw_file_path)
# ファイルを開いて出力
with open(heading1_file_path, 'w', encoding='utf-8') as heading1_file, \
     open(heading2_file_path, 'w', encoding='utf-8') as heading2_file:

    first_heading1 = True  # 最初の見出し1を処理するフラグ

    # Word文書内のすべての段落をループ
    for paragraph in doc.Paragraphs:
        word_range = paragraph.Range
        
        # 見出し1のチェック
        if is_heading1(word_range):
            heading1_file.write(word_range.Text.strip() + '\n')
            if not first_heading1:
                heading2_file.write('\n')  # 最初の見出し1以降で改行を追加
            first_heading1 = False
        
        # 見出し2のチェック
        elif is_heading2(word_range):
            heading2_file.write(word_range.Text.strip() + '\n')

# Word文書を閉じる
doc.Close(False)
word.Quit()

print(f"見出し1のテキストは {heading1_file_path} に保存されました。")
print(f"見出し2のテキストは {heading2_file_path} に保存されました。")
