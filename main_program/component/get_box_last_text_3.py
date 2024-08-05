import win32com.client
import re

def remove_duplicate_numbers_with_ret(text):
    # 連続する重複した数字を1回のみ表示する
    # ただし、末尾の "\r" は除外する
    cleaned_text = ''
    prev_char = ''
    for char in text:
        if char.isdigit() and char == prev_char:
            continue
        cleaned_text += char
        prev_char = char
    # 末尾の "\r" を追加
    if text.endswith('\r'):
        cleaned_text += '\r'
    return cleaned_text.strip()

def clean_text(text):
    # \uXXXX のようなコード部分を取り除く
    return re.sub(r'\\u[0-9A-Fa-f]{4}', '', text)

def last_text(docx_file):
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False  # Wordアプリケーションを非表示にする
    doc = word.Documents.Open(docx_file)
    box_last_text = []
    capture_next = False

    for story_range in doc.StoryRanges:
        for paragraph in story_range.Paragraphs:
            text = paragraph.Range.Text
            if not text:
                continue
            # デバッグ情報を表示
            print(f"[DEBUG] Text: {repr(text)}")

            if capture_next:
                if text.strip() == '\x07':
                    cleaned_text = clean_text(last_text_segment)
                    box_last_text.append(cleaned_text)
                    capture_next = False

            if '\r\x07' in text:
                last_text_segment = text.replace('\r\x07', '').strip()
                capture_next = True

    doc.Close()
    word.Quit()

    return box_last_text

docx_file_path = r'../input/240725_2.docx'
box_last_text = last_text(docx_file_path)

# テキストをファイルに書き出し
output_file_path = r'../output/box_last_text_output_3.txt'
with open(output_file_path, 'w', encoding='utf-8') as f:
    for text in box_last_text:
        text = remove_duplicate_numbers_with_ret(text)
        cleaned_text = clean_text(text)
        text = text.replace('　','')
        print(text)
        f.write(text + '\n')

print("テキストがファイルに出力されました:", output_file_path)
