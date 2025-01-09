import win32com.client
import os
import re

# Wordアプリケーションを起動
word = win32com.client.Dispatch("Word.Application")
word.Visible = False  # Wordアプリケーションを非表示にする

# Word文書を開く
script_directory = os.path.dirname(os.path.abspath(__file__))
docx_file_path = os.path.abspath(os.path.join(script_directory, '..', 'input', '240527_2.docx'))

doc = word.Documents.Open(docx_file_path)

# 出力ファイルのパスを指定
output_file_path = os.path.abspath(os.path.join(script_directory, '..', 'output', 'comments_output_05_2.txt'))

# キーワードの指定
keywords = ["遷移"]

# URLを抽出するための正規表現パターン
url_pattern = re.compile(r'https://[^\s]+')

# ハイパーリンク後のテキストを抽出するための正規表現パターン
hyperlink_text_pattern = re.compile(r'\"”_blank”\"(.+)')

# コメントをテキストファイルに保存
with open(output_file_path, 'w', encoding='utf-8') as file:
    for comment in doc.Comments:
        # コメントが紐づいているテキスト範囲
        comment_range = comment.Scope.Text.strip()  # 余分な空白を除去

        # コメントのテキスト
        comment_text = comment.Range.Text.strip()  # 余分な空白を除去

        # ハイパーリンクのテキスト部分を抽出
        hyperlink_match = hyperlink_text_pattern.search(comment_range)
        if hyperlink_match:
            comment_range = hyperlink_match.group(1).strip()

        # 指定されたキーワードがコメント内容に含まれているかチェック
        if any(keyword in comment_text for keyword in keywords):
            comment_text = comment_text.replace('','').replace('','').replace('','').replace('\x07','')
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

# Word文書を閉じる
doc.Close(False)
word.Quit()

print(f"コメントが保存されました: {output_file_path}")
