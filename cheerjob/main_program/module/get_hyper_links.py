from lxml import etree
import zipfile
import os
import json
import argparse
import re

def extract_hyperlink_texts(docx_path):
    # Word文書の中のリレーションファイル（XML）を抽出
    with zipfile.ZipFile(docx_path, 'r') as docx:
        # Word文書の内容を取得
        xml_content = docx.read('word/document.xml')
        tree = etree.fromstring(xml_content)

    # ハイパーリンクのテキストを格納するセット（重複排除）
    hyperlink_texts = set()

    # ハイパーリンクの要素を検索
    for elem in tree.iter():
        if elem.tag.endswith('hyperlink'):
            # ハイパーリンクのテキストを取得
            text = ''.join([node.text for node in elem.iter() if node.text])
            if text:
                hyperlink_texts.add(text)  # 重複を防ぐためセットに追加

    return hyperlink_texts

def save_hyperlink_texts_to_file(hyperlink_texts, output_file_path):
    # ファイルに出力（ハイパーリンクのテキストがない場合でもファイルを作成）
    with open(output_file_path, 'w', encoding='utf-8') as file:
        if hyperlink_texts:
            for text in hyperlink_texts:
                file.write(f"{text}\n")
            print(f"ハイパーリンクのテキストがファイルに保存されました: {output_file_path}")
        else:
            print("ハイパーリンクテキストはありませんでした。")
            file.write("")  # 空のファイルを作成

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
hyper_links_file_path = os.path.abspath(os.path.join(base_dir, data["hyper_links_file_path"]))

# docx_file_path_2 を生成
docx_raw_file_name = os.path.basename(docx_raw_file_path)
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
hyperlink_texts = extract_hyperlink_texts(docx_file_path_2)
save_hyperlink_texts_to_file(hyperlink_texts, hyper_links_file_path)
