# import chardet
import re
import os

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

def process_text_file(input_file_path):
    with open(input_file_path, 'r', encoding='utf-8') as file:
        content = file.read()
    
    # テキストを2行連続の改行で分割
    blocks = content.split('\n\n')

    results = []
    
    for block in blocks:
        # 前後の空白を削除
        block = block.strip()
        
        # 空のブロックはスキップ
        if not block:
            continue
        
        # ブロックを行に分割
        lines = block.split('\n')
        
        # "目次" を含む行を削除
        filtered_lines = [line for line in lines if "目次" not in line]
        
        # フィルタリング後に空のブロックはスキップ
        if not filtered_lines:
            continue
        
        # フィルタリングされたブロックの最後の行を取得
        last_line = filtered_lines[-1].strip()
        if last_line:  # 空でないことを確認
            results.append(last_line)
    
    return results

# 使用例
script_directory = os.path.dirname(os.path.abspath(__file__))
input_file_path = os.path.abspath(os.path.join(script_directory, '..','..', 'output', 'get_border_text_08_1.html'))

# 関数を呼び出して結果を配列に格納
result_array = process_text_file(input_file_path)

# 結果をファイルに書き出し
# output_file_path = r'../output/result_array_output.txt'
# with open(output_file_path, 'w', encoding='utf-8') as f:
for text in result_array:
    # テキストを加工
    text = remove_duplicate_numbers_with_ret(text)
    cleaned_text = clean_text(text)
    text = text.replace(' ', '')  # スペースを削除
    
    # repr を使って内容を確認するために出力
    # print(text)
    
    # # ファイルに書き出し
    # f.write(text + '\n')

# print("テキストがファイルに出力されました:", output_file_path)
