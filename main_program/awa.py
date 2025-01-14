﻿import os
import colorsys
import win32com.client as win32
import re
from module.create_table_with_color import create_html_table
from module import border_last_text
from module import border_text
def remove_digits_from_word_range(word_range):
    word_text_without_digits = ''
    if word_range is None:
        return word_text_without_digits

    try:
        # 現在のword_rangeのテキストを取得
        word_text = word_range.Text.strip()

        # 数字部分を削除するための正規表現
        word_text_without_digits = re.sub(r'\d+', '', word_text)

        return word_text_without_digits

    except Exception as e:
        print(f"Error removing digits from word range: {e}")
        return word_text_without_digits
def check_prev_digits_in_word_range(word_range, prev_word_range):
    if word_range is None or prev_word_range is None:
        return False  # Noneが渡された場合は、Falseを返す

    prev_text = prev_word_range.Text.strip()
    prev_digits = ''.join(filter(str.isdigit, prev_text))

    if prev_digits:
        try:
            # 現在のword_rangeのテキストを取得
            word_text = word_range.Text.strip()

            # prev_digitsがword_textに含まれているか確認
            if prev_digits in word_text:
                return True  # prev_digitsがword_textに含まれている場合はTrueを返す
            else:
                return False  # 含まれていない場合はFalseを返す

        except Exception as e:
            print(f"Error checking word range: {e}")
            return False  # エラーが発生した場合はFalseを返す

    return False  # prev_digitsが存在しない場合はFalseを返す
def modify_word_range_text(word_range, prev_word_range):
    if prev_word_range is not None:
        prev_text = prev_word_range.Text.strip()

        # 数字の羅列を見つけるための正規表現パターン
        number_pattern = r'\d+'

        # prev_word_range から数字の羅列をすべて抽出
        prev_numbers = re.findall(number_pattern, prev_text)

        # それぞれの数字の羅列が現在のテキストに含まれている場合は削除
        text = word_range.Text
        for number in prev_numbers:
            if number in text:
                print(f"number:{number},text:{text}")
                # 数字の羅列を削除
                text = text.replace(number, '')
                # word_rangeのテキストを更新
                word_range.Text = text
    
    return word_range

def format_text_block_to_html(text_block):
    """テキストブロックをHTMLに変換する関数。各行が<p>タグ内の文字列と一致する場合にのみ処理を実行"""
    # <p>タグの内容を取得
    paragraph_matches = re.findall(r'<p>(.*?)</p>', text_block, re.DOTALL)
    
    if not paragraph_matches:
        return text_block  # 一致しない場合は元のテキストを返す

    # ブロック内のすべての行が<p>タグ内の文字列と一致するか確認
    lines = text_block.strip().split('\n')
    if all(any(line.strip() in paragraph for paragraph in paragraph_matches) for line in lines):
        # 全ての行が<p>タグ内の文字列と一致する場合、<ul><li>で囲う
        html_output = '<div style="background:#ffffff;border:1px solid #cccccc;padding:5px 10px;">\n<ul>\n'
        for line in lines:
            line_content = line.lstrip('・').strip()
            html_output += f'<li><strong>{line_content}</strong></li>\n'
        html_output += '</ul>\n</div>\n'
    else:
        # 一致しない場合、各行をそのまま出力し、<br>で区切る
        html_output = '<div style="background:#ffffff;border:1px solid #cccccc;padding:5px 10px;">\n'
        for line in lines:
            html_output += f'<strong>{line.strip()}</strong><br />\n'
        html_output += '</div>\n'

    return html_output

def read_html_tables(html_file_path):
    """HTMLファイルを読み込み、テーブル部分をリストとして返す関数"""
    with open(html_file_path, 'r', encoding='utf-8') as file:
        content = file.read()

    # テーブル部分を抽出してリスト化
    tables = re.findall(r'<table.*?</table>', content, re.DOTALL)
    return tables

def remove_duplicate_english(text1, text2):
    # 英数字部分を正規表現で検出
    pattern = r'[a-zA-Z0-9]+'
    
    match1 = re.search(pattern, text1)
    match2 = re.search(pattern, text2)
    
    if match1 and match2:
        # 両方に英数字が存在し、かつそれが同じ場合に後者の英数字部分を削除
        if match1.group() == match2.group():
            text1 = text1.replace(match2.group(), '', 1)
    
    return text1

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

def remove_trailing_digits(text):
    # 文字列の末尾が数字である限り、末尾の数字を取り除く
    while text and text[-1].isdigit():
        text = text[:-1]
    return text.strip()
def make_list_subtitle(text):
    text = text.replace('<span class="marker">','').replace('<p>','').replace('\x07','')
    # Split the text by newlines and remove any leading or trailing whitespace
    lines = text.strip().split('\r')

    # Remove the leading '・' from each line and store them in a list
    result = [line.lstrip('・') for line in lines]
    print(result)
    return result

def is_heading1(word_range):
    if word_range is None:
        return False  # NoneであればFalseを返す
    if word_range.Style is not None:
        # スタイルが見出し1であるかどうかのチェック
        return word_range.Style.NameLocal == "見出し 1"
    return False

def is_heading2(word_range):
    if word_range is None:
        return False  # NoneであればFalseを返す
    # 指定されたWordのRangeが「見出し2」のスタイルかどうかを判定する関数。
    # word_range.StyleがNoneでないことを確認
    if word_range.Style is not None:
        return word_range.Style.NameLocal == "見出し 2"
    return False

def rgb_to_hsv(rgb):
    # RGB形式からHSV形式に変換する関数
    r, g, b = rgb
    h, s, v = colorsys.rgb_to_hsv(r / 255, g / 255, b / 255)
    return (h * 360, s * 100, v * 100)

def is_blue_color(word_range):
    # 単語の色が青色かどうかをチェックします
    # RGB形式に変換してから比較します
    color_value = word_range.Font.Color
    blue_rgb = (color_value & 0xFF, (color_value >> 8) & 0xFF, (color_value >> 16) & 0xFF)
    blue_hsv = rgb_to_hsv(blue_rgb)
    return blue_hsv[0] >= 200 and blue_hsv[0] <= 260  # 青色のHSV範囲（Hue）

def is_yellow_color(word_range):
    if word_range.HighlightColorIndex == 7:  # 7は黄色を表す定数
        return True
    return False

def is_end(word_range):
    if word_range.Text.endswith('\n'):
        return True
    elif word_range.Text.endswith('\n\n'):
        return True
    elif word_range.Text.endswith('\r\n'):
        return True
    elif word_range.Text.endswith('\r'):
        return True
    elif word_range.Text == '\r':
        return True
    else:
        return False
    
def check_tag(prev_word_range, word_range, next_word_range, paragraph_text, normal_text, bold_text, blue_text, highlighted_text, box_text, box_count,  box_last_text, last_text_count, prev_is_normal, next_is_table, smart_text):
    if is_blue_color(word_range):
        # 青色の開始
        # bold_text=remove_duplicate_numbers_with_ret(bold_text)
        paragraph_text += normal_text
        normal_text = ''
        if is_end(prev_word_range) or (not is_blue_color(prev_word_range) or (is_heading1(prev_word_range) or is_heading2(prev_word_range))):#前回が太字でない場合
            smart_text.replace('\r', '\n')
            if not prev_is_normal:
                blue_text += f'<p><a href="">{smart_text}'
                prev_is_normal = True
            else:
                if (not is_end(prev_word_range) and prev_word_range.Bold) and not is_blue_color(prev_word_range):
                    bold_text += '</strong>'
                    paragraph_text += bold_text
                    bold_text = ''
                if prev_word_range.Bold:
                    bold_text = ''
                blue_text += f'<a href="">{smart_text}'
        else:
            if prev_is_normal and is_end(word_range) and is_blue_color(next_word_range):
                blue_text += f'</a><br />'
            blue_text += smart_text
        if prev_is_normal and not is_blue_color(next_word_range) or next_is_table:
            if last_text_count < len(box_last_text) and f'{box_last_text[last_text_count]}' in blue_text:#{smart_text}がない
                blue_text = blue_text.replace('<p>', '<div style="background:#ffffff;border:1px solid #cccccc;padding:5px 10px;">') + '</a></div>\r'
                blue_text = blue_text.replace('<br />\r</a></div>','\r</a></div>').replace('<a href="">\r<a href="">','<a href="">').replace('<a href="">\n<a href="">','<a href="">')
                if '・' in blue_text:
                    blue_text = blue_text.replace('<a href="">・','<li><a href="">').replace('</a>','</a></li>').replace('<br />','')
                last_text_count += 1
                prev_is_normal = False
            else:
                smart_text.replace('\r', '\n')
                if is_end(next_word_range):
                    blue_text += f'</a></p>\r'
                    prev_is_normal = False
                else:
                    if '。'in f"{smart_text}":
                        blue_text += f'</a>{smart_text}</p>\r'
                        prev_is_normal = False
                    else:
                        blue_text += f'</a>'
            paragraph_text += blue_text
            blue_text = ''
            normal_text = ''
    # ハイライトの開始
    elif is_yellow_color(word_range):
        # bold_text=remove_duplicate_numbers_with_ret(bold_text)
        paragraph_text += normal_text
        normal_text = ''
        if (not is_end(prev_word_range) and prev_word_range.Bold) and not is_yellow_color(prev_word_range):
            bold_text += '</strong>'
            paragraph_text += bold_text
            bold_text = ''
        if is_end(prev_word_range) or (not is_yellow_color(prev_word_range)) or is_heading1(prev_word_range) or is_heading2(prev_word_range):#前回が太字でない場合
            smart_text.replace('\r', '\n')
            if not prev_is_normal:
                highlighted_text += f'<p><span class="marker"><strong>{smart_text}'
                prev_is_normal = True
            else:
                if prev_word_range.Bold:
                    bold_text = ''
                highlighted_text += f'<span class="marker"><strong>{smart_text}'
        else:
            highlighted_text += smart_text

        if prev_is_normal and ((next_word_range is not None and not is_yellow_color(next_word_range)) and not is_heading1(next_word_range) and not is_heading2(next_word_range) or is_end(next_word_range)):
            highlighted_text = highlighted_text.replace('<span class="marker"><strong></strong></span>','').replace('</strong></span><span class="marker"><strong>','')
            # print(highlighted_text)
            if last_text_count < len(box_last_text) and f'{box_last_text[last_text_count]}' in highlighted_text:#{smart_text}がない
                highlighted_text = highlighted_text.replace('<p>', '<div style="background:#ffffff;border:1px solid #cccccc;padding:5px 10px;">') + '</strong></div>\r'
                highlighted_text = highlighted_text.replace('<br />\r</strong></span></div>','\r</strong></span></div>')
                if '・' in highlighted_text:
                    highlighted_text = highlighted_text.replace('<span class="marker"><strong>・','<li><span class="marker"><strong>').replace('</strong></span>','</strong></span></li>').replace('<br />','')
                highlighted_text = highlighted_text.replace('<br />','<br />\r')
                last_text_count += 1
                box_count = 0
                prev_is_normal = False
                # print(highlighted_text)
                paragraph_text += normal_text + highlighted_text
                highlighted_text = ''
                normal_text = ''
            else:
                if is_end(next_word_range) or is_heading1(next_word_range) or is_heading2(next_word_range):
                    if last_text_count < len(box_text) and box_count < len(box_text[last_text_count]) and f'{box_text[last_text_count][box_count]}' in highlighted_text:
                        highlighted_text += f'</strong></span><br />\r'
                        box_count += 1
                    else:
                        highlighted_text += f'</strong></span></p>\r'
                        prev_is_normal = False
                        paragraph_text += normal_text + highlighted_text
                        highlighted_text = ''
                        normal_text = ''
                else:
                    highlighted_text += f'</strong></span>'#ここで<p></strong>が発生している
                    paragraph_text += normal_text + highlighted_text
                    highlighted_text = ''
                    normal_text = ''
    # 太字の開始
    elif word_range.Bold:
        paragraph_text += normal_text
        normal_text = ''
        # bold_text=remove_duplicate_numbers_with_ret(bold_text)
        if prev_word_range is not None:
            if is_end(prev_word_range) or not prev_word_range.Bold or is_yellow_color(prev_word_range) or is_heading1(prev_word_range) or is_heading2(prev_word_range):#前回が太字でない場合
                if box_count == 0:
                    bold_text = ''
                if not prev_is_normal:#テキストが始まってもない場合
                    if is_end(next_word_range) or is_end(word_range):
                        bold_text += f'<p><strong>{smart_text}</strong></p>'
                        paragraph_text += normal_text + bold_text
                        bold_text = ''
                        normal_text = ''
                    else:
                        bold_text += f'<p><strong>{smart_text}'
                        prev_is_normal = True
                else:
                    bold_text += f'<strong>{smart_text}'
            else:
                bold_text += smart_text
            if prev_is_normal and ((next_word_range is not None and not next_word_range.Bold) and not is_heading1(next_word_range) and not is_heading2(next_word_range) or is_end(next_word_range)):
                bold_text = bold_text.replace('<strong></strong>','').replace('</strong><strong>','')
                # print(bold_text)
                if last_text_count < len(box_last_text) and f'{box_last_text[last_text_count]}' in bold_text:#{smart_text}がない
                    bold_text = bold_text.replace('<p>', '<div style="background:#ffffff;border:1px solid #cccccc;padding:5px 10px;">') + '</strong></div>\r'
                    bold_text = bold_text.replace('<br />\r</strong></div>','\r</strong></div>')
                    if '・' in bold_text:
                        bold_text = bold_text.replace('<strong>・','<li><strong>').replace('</strong>','</strong></li>').replace('<br />','')
                    bold_text = bold_text.replace('<br />','<br />\r').replace('\r\r','\r')
                    last_text_count += 1
                    box_count = 0
                    prev_is_normal = False
                    # print(bold_text)
                    paragraph_text += normal_text + bold_text
                    bold_text = ''
                    normal_text = ''
                else:
                    if is_end(next_word_range) or is_heading1(next_word_range) or is_heading2(next_word_range):
                        if last_text_count < len(box_text) and box_count < len(box_text[last_text_count]) and f'{box_text[last_text_count][box_count]}' in bold_text:
                            bold_text += f'</strong><br />\r'
                            box_count += 1
                        else:
                            bold_text += f'</strong></p>\r'
                            prev_is_normal = False
                            paragraph_text += normal_text + bold_text
                            bold_text = ''
                            normal_text = ''
                    else:
                        bold_text += f'</strong>'#ここで<p></strong>が発生している
                        paragraph_text += normal_text + bold_text
                        bold_text = ''
                        normal_text = ''
    # マーカーや青色のテキスト以外のテキスト
    elif next_word_range is not None and is_end(next_word_range)  or '。'in f"{smart_text}":
        normal_text += f'{smart_text}'
        if not prev_is_normal:
            paragraph_text += f'<p>{smart_text}</p>'
        else:
            if last_text_count < len(box_last_text) and f'{box_last_text[last_text_count]}' in normal_text:#{smart_text}がない
                    normal_text = normal_text.replace('\n','\r')
                    normal_text = normal_text.replace('<p>', '<div style="background:#ffffff;border:1px solid #cccccc;padding:5px 10px;">') + '</div>\r'
                    normal_text = normal_text.replace('<br />\r</div>','\r</div>')
                    if '・' in normal_text:
                        normal_text = normal_text.replace('\r・','\r<li>').replace('<br />','</li>').replace('\r</div>','</li>\r</div>').replace('<div style="background:#ffffff;border:1px solid #cccccc;padding:5px 10px;">・','<div style="background:#ffffff;border:1px solid #cccccc;padding:5px 10px;">\r<li>')
                    last_text_count += 1
                    box_count = 0
                    prev_is_normal = False
                    print(normal_text)
                    paragraph_text += normal_text
                    normal_text = ''
            elif last_text_count < len(box_text) and box_count < len(box_text[last_text_count]) and f'{box_text[last_text_count][box_count]}' in normal_text:
                normal_text += f'<br />'
                box_count += 1
            else:
                normal_text += f'</p>\r'
                paragraph_text += normal_text
                normal_text = ''
                prev_is_normal = False
    elif prev_is_normal: #次で終わりでないが、テキストはすでに始まっている場合
        if f'<p>▼関連記事はこちら' in normal_text:
            normal_text=f'<p>▼関連記事はこちら<br />\r'
        else:
            normal_text += f'{smart_text}'
        # prev_is_normal = True
    else:#テキストが始まっていない場合
        if not is_end(word_range):
            normal_text += f'<p>{smart_text}'
            prev_is_normal = True
    # else:
    #     paragraph_text += f'{smart_text}'#なぜか数十個の\rが表示される,本来is_end()で引っかかるはず
    return paragraph_text, normal_text, bold_text, blue_text, highlighted_text, box_text, box_count, box_last_text, last_text_count,prev_is_normal

def extract_text_with_markup(docx_file, html_tables):
    script_directory = os.path.dirname(os.path.abspath(__file__))
    docx_file_path = os.path.abspath(os.path.join(script_directory, '..', 'input', '240821_1.docx'))
    input_file_path = os.path.abspath(os.path.join(script_directory, '..', 'output', 'get_border_text_08_1.html'))
    print(f"Trying to open: {docx_file_path}")

    # テスト用にファイルの存在確認
    if not os.path.exists(docx_file_path):
        raise FileNotFoundError(f"File not found: {docx_file_path}")
    
    # テキストファイルの内容を読み込み、リストに保存
    if not os.path.exists(input_file_path):
        raise FileNotFoundError(f"File not found: {input_file_path}")
    
    box_text = border_text.txt_to_2d_array(input_file_path)
    # print(f"box_text[0][0]:{box_text[0][0]}")
    box_count = 0

    #箱の最後のテキストを取り出す
    box_last_text = border_last_text.process_text_file(input_file_path)
    # print(f"box_last_text[0]:{box_last_text[0]}")
    last_text_count = 0
    print(1)

    word = win32.DispatchEx("Word.Application")
    word.Visible = False  # Wordアプリケーションを非表示にする

    doc = word.Documents.Open(docx_file)
    extracted_text = []

    # 青色のテキストを一時的に保持する変数
    blue_text = ''
    # 太字のテキストを一時的に保持する変数
    bold_text = ''
    # 黄色のマーカーのテキストを一時的に保持する変数
    highlighted_text = ''
    # 見出し1のテキストを一時的に保持する変数
    h3_text = ''
    # 見出し2のテキストを一時的に保持する変数
    h4_text = ''
    # 普通のテキストを一時的に保持する変数
    normal_text = ''
    smart_text = ''
    prev_word_range = None
    next_word_range = None
    prev_is_normal = False
    word_is_duplicate = False
    next_is_table = False
    in_table = False
    wdInTable = 12

    # HTMLテーブルをリストとして管理
    table_index = 0

    for range in doc.StoryRanges:
        # 各段落内のテキストを結合して1つの文にする
        word_ranges = list(range.Words)  # Wordsをリストに変換
        paragraph_text = ''
        for i, word_range in enumerate(word_ranges):
            next_word_range = None
            # word_text = word_range.Text
            prev_word_range = word_ranges[i - 1] if i > 0 else None
            if (i + 1 < len(word_ranges)):
                next_word_range = word_ranges[i + 1]
            if next_word_range is not None:
                next_is_table = next_word_range.Information(wdInTable)
            else:
                next_is_table = False
            if not word_range.Information(wdInTable):#表のテキストでない
                in_table = False#表の初めのテキストでない
                word_is_duplicate = check_prev_digits_in_word_range(word_range, prev_word_range)
                if word_is_duplicate:
                    smart_text = remove_digits_from_word_range(word_range)
                else:
                    smart_text = f'{word_range.Text}'
                if is_heading1(word_range):
                    normal_text = ''
                    bold_text = ''
                    if not is_heading1(prev_word_range):
                        paragraph_text += f'<h3>'
                    # 見出し1スタイルのテキストである場合の処理
                    h3_text += f"{word_range.Text.strip()}"
                    if next_word_range is not None and not is_heading1(next_word_range) or next_is_table:
                        h3_text += f"</h3>\r"
                        paragraph_text += h3_text
                        h3_text = ''
                        paragraph_text, normal_text, bold_text, blue_text, highlighted_text, box_text, box_count, box_last_text, last_text_count, prev_is_normal = check_tag(
            prev_word_range, word_range, next_word_range, paragraph_text, normal_text, bold_text, blue_text, highlighted_text, box_text, box_count, box_last_text,last_text_count, prev_is_normal, next_is_table, smart_text)
                elif is_heading2(word_range):
                    normal_text = ''
                    bold_text = ''
                    if not is_heading2(prev_word_range):
                        paragraph_text += f'<h4>'
                    # 見出し1スタイルのテキストである場合の処理
                    h4_text += f"{word_range.Text.strip()}"
                    if next_word_range is not None and not is_heading2(next_word_range) or next_is_table:
                        h4_text += f"</h4>\r"
                        paragraph_text += h4_text
                        h4_text = ''
                        paragraph_text, normal_text, bold_text, blue_text, highlighted_text, box_text, box_count, box_last_text,last_text_count, prev_is_normal = check_tag(
            prev_word_range, word_range, next_word_range, paragraph_text, normal_text, bold_text, blue_text, highlighted_text, box_text, box_count, box_last_text, last_text_count, prev_is_normal, next_is_table, smart_text)
                else:
                    paragraph_text, normal_text, bold_text, blue_text, highlighted_text, box_text, box_count, box_last_text,last_text_count, prev_is_normal = check_tag(
            prev_word_range, word_range, next_word_range, paragraph_text, normal_text, bold_text, blue_text, highlighted_text, box_text, box_count, box_last_text, last_text_count, prev_is_normal, next_is_table, smart_text)
            else:
                if not in_table:
                    print(2)
                     # リストにあるHTMLテーブルを順番に追加
                    if table_index < len(html_tables):
                        paragraph_text += html_tables[table_index]
                        table_index += 1
                        print(3)
                    in_table = True
                
        # テキストが空でない場合のみ処理を行います
        if paragraph_text:
            paragraph_text = paragraph_text.replace('<p>/</p>','').replace('<strong>\r</strong>','').replace('<p></p>','').replace('<strong></strong>','').replace('</strong><strong>','').replace('<p>\r</p>','').replace('<p>\r','<p>').replace('','').replace('','').replace('','').replace('<p>▼関連記事はこちら</p>\r<p>','<p>▼関連記事はこちら<br />\r').replace('<p>/</p>','').replace('<br />\r</span></div>','\r</span></div>').replace('<p> </p>','').replace('</a></p>\r<p><a href="">',',</a><br />\r<a href="">').replace('</a></p>\n<p><a href="">',',</a><br />\r<a href="">').replace('</a></p>\n<p>',',</a><br />\r').replace('▼関連記事はコチラ</p>\r<p><a href="">','▼関連記事はコチラ<br />\r<a href="">').replace('▼関連記事はコチラ</p>\n<p><a href="">','▼関連記事はコチラ<br />\n<a href="">')
            
            # 文の末尾に数字がある場合、その数字を取り除く
            cleaned_text = remove_trailing_digits(paragraph_text)
            # 連続する数字を1回のみ表示する
            # cleaned_text = remove_duplicate_numbers_with_ret(cleaned_text)
            if cleaned_text:
                extracted_text.append(cleaned_text)
        break

    doc.Close()
    word.Quit()
    return extracted_text

script_directory = os.path.dirname(os.path.abspath(__file__))
docx_file_path = os.path.abspath(os.path.join(script_directory, '..', 'input', '240821_1.docx'))
html_file_path = os.path.abspath(os.path.join(script_directory, '..', 'output', 'combined_tables_08_1.html'))


# HTMLファイルからテーブルを読み込み、リストとして管理
html_tables = read_html_tables(html_file_path)

# WordファイルのテキストとHTMLテーブルを統合して抽出
extracted_text_with_markup = extract_text_with_markup(docx_file_path, html_tables)

print(99)
html_output = ''.join(extracted_text_with_markup)
output_file_path = os.path.abspath(os.path.join(script_directory, '..', 'output', 'extracted_text_knowledge_08_1n.html'))
with open(output_file_path, 'w', encoding='utf-8') as html_file:
    html_file.write(html_output)
print(100)

print(f'HTMLファイル "{output_file_path}" に抽出されたテキストが保存されました。')