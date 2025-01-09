import json
import argparse
import re
import os
import win32com.client as win32

def remove_comments_from_word(doc_path, save_path=None):
    # Wordアプリケーションを起動
    word = win32.Dispatch("Word.Application")
    word.Visible = False

    # 指定されたWord文書を開く
    doc = word.Documents.Open(doc_path)

    # コメントを1つずつ削除
    for comment in doc.Comments:
        comment.Delete()

    # 上書き保存または別名保存
    if save_path:
        doc.SaveAs2(save_path)
    else:
        doc.Save()

    # Word文書を閉じる
    doc.Close()
    word.Quit()
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
hyperlink_file_path = os.path.abspath(os.path.join(base_dir, data["hyperlink_file_path"]))
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
docx_file_name_modified = docx_raw_file_name.replace('.docx', '_without_toc_final_no_images.docx')
docx_file_path_2 = os.path.abspath(os.path.join(output_dir, docx_file_name_modified))

# ファイルパスを指定
docx_remove_hyperlinks_path = docx_raw_file_name.replace('.docx', '_without_toc_final_no_images_remove_hyperlinks.docx')
docx_remove_comments_path = docx_raw_file_name.replace('.docx', '_without_toc_final_no_images_remove_hyperlinks_remove_comments.docx')
input_docx_file_path = os.path.abspath(os.path.join(output_dir, docx_remove_hyperlinks_path))
output_docx_file_path = os.path.abspath(os.path.join(output_dir, docx_remove_comments_path))

# コメントを削除
remove_comments_from_word(input_docx_file_path, output_docx_file_path)
