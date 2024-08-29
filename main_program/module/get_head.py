import win32com.client
import os

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

# Word文書を開く
script_directory = os.path.dirname(os.path.abspath(__file__))
docx_file_path = os.path.abspath(os.path.join(script_directory, '..', '..', 'input', '240418【校了】自己分析_看護師_転職成功.docx'))

doc = word.Documents.Open(docx_file_path)

# 出力ファイルのパスを設定
heading1_file_path = os.path.abspath(os.path.join(script_directory, '..', '..', 'output', 'heading4_1.txt'))
heading2_file_path = os.path.abspath(os.path.join(script_directory, '..', '..', 'output', 'heading4_2.txt'))

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
