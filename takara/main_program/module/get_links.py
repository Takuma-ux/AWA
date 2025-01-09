import win32com.client
import os
import re
import json
import argparse

# Wordアプリケーションを起動
word = win32com.client.Dispatch("Word.Application")
word.Visible = False  # Wordアプリケーションを非表示にする

# Word文書を開く
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
links_file_path = os.path.abspath(os.path.join(base_dir, data["links_file_path"]))

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

if not os.path.exists(docx_file_path_2):
    print("File does not exist:", docx_file_path_2)
else:
    doc = word.Documents.Open(docx_file_path_2)

# キーワードの指定
keywords = ["遷移"]

# URLを抽出するための正規表現パターン
url_pattern = re.compile(r'https://[^\s]+')

# ハイパーリンク後のテキストを抽出するための正規表現パターン
hyperlink_text_pattern = re.compile(r'\"”_blank”\"(.+)')

# コメントをテキストファイルに保存
with open(links_file_path, 'w', encoding='utf-8') as file:
    for comment in doc.Comments:
        # コメントが紐づいているテキスト範囲
        comment_range = comment.Scope.Text.strip()  # 余分な空白を除去
        comment_range = comment_range.replace("\n", "").replace('\r','')  # 改行を削除

        # コメントのテキスト
        comment_text = comment.Range.Text.strip()  # 余分な空白を除去
        comment_text = comment_text.replace("\n", "").replace('\r','')  # 改行を削除

        # ハイパーリンクのテキスト部分を抽出
        hyperlink_match = hyperlink_text_pattern.search(comment_range)
        if hyperlink_match:
            comment_range = hyperlink_match.group(1).strip()

        # 指定されたキーワードがコメント内容に含まれているかチェック
        if any(keyword in comment_text for keyword in keywords):
            comment_text = comment_text.replace('','').replace('','').replace('','').replace('\x07','')
            comment_text = comment_text.replace("\n", "").replace('\r','')  # 改行を削除
            # URLの抽出
            url_match = url_pattern.search(comment_text)
            if url_match:
                comment_text = url_match.group(0)  # 最初に見つかったURLを取得

            # "大見出し"または"小見出し"の抽出
            heading_match = re.search(r'(大見出し|小見出し)「[^」]+」', comment_text)
            if heading_match:
                comment_text = heading_match.group(0)  # "大見出し"または"小見出し"とそれに続くテキストを取得
            
            # 有効な情報が含まれているかをチェック（URL、小見出し、大見出しのいずれか）
            if url_match or heading_match:
                # ファイルに書き込む形式: [コメント対象のテキスト],[コメント内容]
                file.write(f"[{comment_range}],[{comment_text}]\n")
            else:
                file.write("")  # 空のファイルを作成


# Word文書を閉じる
doc.Close(False)
word.Quit()

print(f"コメントが保存されました: {links_file_path}")
