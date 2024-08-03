import win32com.client
import re

def extract_text_from_docx(docx_file_path):
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False

    doc = word.Documents.Open(docx_file_path)
    extracted_text = []

    for paragraph in doc.Paragraphs:
        extracted_text.append(paragraph.Range.Text.strip())

    doc.Close()
    word.Quit()

    return extracted_text

def extract_lines_ending_with_number(text):
    pattern = r'.*\d$'
    matches = [line for line in text if re.match(pattern, line)]
    return matches

def remove_numbers_from_end(text):
    cleaned_lines = [re.sub(r'\d+$', '', line) for line in text]
    return cleaned_lines

# Wordファイルのパス
docx_file_path = r'C:\Users\takum\Documents\cheerjob_auto\test_0802.docx'

# Wordファイルからテキストを抽出
extracted_text = extract_text_from_docx(docx_file_path)

# 数字で終わる行を抽出
lines_ending_with_number = extract_lines_ending_with_number(extracted_text)

# 末尾の数字を削除して各行のテキストを取り出す
cleaned_text = remove_numbers_from_end(lines_ending_with_number)

# 末尾の数字を削除して各行のテキストを取り出す
cleaned_text = [line.rstrip() for line in cleaned_text]