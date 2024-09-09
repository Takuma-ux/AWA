import os
import colorsys
import win32com.client as win32
import re
import difflib
import json
import argparse
from module import border_last_text
from module import border_text
from module import get_hyper_link_text

def remove_duplicate_english_words(word_range,prev_word_range,smart_text):
    word_text_without_duplicates = ''
    if word_range is None:
        return word_text_without_duplicates

    try:
        # 現在のword_rangeのテキストを取得
        word_text = smart_text
        word_text_1 = prev_word_range.Text.strip()

        # 英語の単語だけを抽出する正規表現
        english_words = re.findall(r'[A-Za-z]+', smart_text)
        english_words_1 = re.findall(r'[A-Za-z]+', word_text_1)

        # 重複をチェックしながらテキストを再構築
        word_text_without_duplicates = word_text
        for word in english_words:
            if word in english_words_1:
                # 重複がある場合、該当する英語の単語を削除
                word_text_without_duplicates = re.sub(r'\b' + re.escape(word) + r'\b', '', word_text_without_duplicates).strip()

        return word_text_without_duplicates

    except Exception as e:
        print(f"Error removing duplicate English words from word range: {e}")
        return word_text_without_duplicates

def remove_surrounding_text(comment_text):
    # 「大見出し」「小見出し」および「」を除去
    comment_text = comment_text.replace("大見出し", "").replace("小見出し", "")
    comment_text = comment_text.replace("「", "").replace("」", "")
    return comment_text.strip()

def is_similar(text1, text2, threshold=0.8):
    # text1とtext2の類似度を計算
    similarity = difflib.SequenceMatcher(None, text1, text2).ratio()
    return similarity >= threshold

def load_headings(heading1_file, heading2_file):
    heading1_list = []
    heading2_list = []

    # 見出し1のファイルを読み込み
    with open(heading1_file, 'r', encoding='utf-8') as file:
        heading1_list = [line.strip() for line in file if line.strip()]  # 空行を除いてリストに追加

    # 見出し2のファイルを読み込み
    with open(heading2_file, 'r', encoding='utf-8') as file:
        heading2_group = []
        for line in file:
            line = line.strip()
            if line == "":  # 空行があった場合
                heading2_list.append(heading2_group)  # 現在のグループを追加
                heading2_group = []  # 次のグループの準備
            else:
                heading2_group.append(line)
        heading2_list.append(heading2_group)  # 最後のグループを追加

    # 見出し1に対応する見出し2を調整
    if len(heading1_list) > len(heading2_list):
        heading2_list.append([])  # 見出し1に対する見出し2がない場合に空のリストを追加

    return heading1_list, heading2_list

def load_comments_from_file(file_path):
    comments_list = []

    # テキストファイルを読み込み
    with open(file_path, 'r', encoding='utf-8') as file:
        for line in file:
            # 行末の改行を削除し、「[」と「]」を取り除く
            line = line.strip().strip('[]')
            
            # 「コメント対象のテキスト」と「コメント内容」を分割
            comment_parts = line.split('],[')
            
            # 「」を取り除く
            cleaned_parts = [part.replace('「', '').replace('」', '') for part in comment_parts]
            
            # 二次元配列としてリストに格納
            comments_list.append(cleaned_parts)

    return comments_list

def process_normal_text(normal_text):
    # 1. "<p>" と "</p>" を除去
    normal_text = normal_text.replace("<p>", "").replace("</p>", "").replace('<br />','')

    # 2. '\r' ごとに分割して normal_split_text に格納
    normal_split_text = normal_text.split('\r')

    # 3. 各項目に対して処理を行う
    processed_texts = []
    for text in normal_split_text:
        stripped_text = text.strip()  # 先頭と末尾の空白を除去

        # 先頭の文字が "・" である場合のみ、それを削除する
        if stripped_text.startswith("・"):
            stripped_text = stripped_text[1:].strip()  # "・" を除去

        # その後、<br /> を追加
        stripped_text = f"{stripped_text}<br />"
        processed_texts.append(stripped_text)

    # 4. 結果を一つの文字列にまとめ、指定のフォーマットで囲む
    final_text = '<div class="solution" style="padding:10px 15px;border:1px solid #000000;">\r' + "\r".join(processed_texts) + '\r</div>'
    final_text = final_text.replace('<br />\r</div>', '\r</div>')
    
    return final_text


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
        html_output = '<div class="solution" style="padding:10px 15px;border:1px solid #000000;">\n<ul>\n'
        for line in lines:
            line_content = line.lstrip('・').strip()
            html_output += f'<li><strong>{line_content}</strong></li>\n'
        html_output += '</ul>\n</div>\n'
    else:
        # 一致しない場合、各行をそのまま出力し、<br>で区切る
        html_output = '<div class="solution" style="padding:10px 15px;border:1px solid #000000;">\n'
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
    # r,g,b=blue_rgb
    # if(r==0 and g==0 and b==256):
    #     return True
    blue_hsv = rgb_to_hsv(blue_rgb)
    # if "NET" in word_range.Text:
    #     print(blue_hsv)
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
    
def check_tag(prev_word_range, word_range, next_word_range, paragraph_text, normal_text, bold_text, blue_text, highlighted_text,hyper_link_text, hyper_text_count, box_text, box_count, box_last_text, last_text_count, comments_list, links_count, heading1_array,heading2_array, h2_count,h3_count, prev_is_normal, next_is_table, smart_text):
    if is_blue_color(word_range):
        # 青色の開始
        # bold_text=remove_duplicate_numbers_with_ret(bold_text)
        paragraph_text += normal_text
        normal_text = ''
        if prev_word_range is not None and (is_end(prev_word_range) or (not is_blue_color(prev_word_range) or (is_heading1(prev_word_range) or is_heading2(prev_word_range)))):#前回が太字でない場合
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
                if word_range.Bold:
                    blue_text += f'<strong><a href="">{smart_text}'
                else:
                    blue_text += f'<a href="">{smart_text}'
        else:
            if prev_is_normal and is_end(word_range) and is_blue_color(next_word_range):
                blue_text += f'</a><br />\r'
                for comment in comments_list:
                    comment_text = comment[0]  # 「コメント対象のテキスト」を取得
                    if f"{comment_text}" in blue_text:
                        # 条件を満たした場合の処理をここに記述
                        if "http" in comment[1]:
                            if "https://www.nurse-step.com/" in comment[1]:
                                link = comment[1]
                                blue_text = blue_text.replace(f'<a href="">{comment[0]}',f'<span style="background:linear-gradient(transparent 60%, #ffe57f 60%);"><a href="{link}">{comment[0]}')
                            else:
                                blue_text = blue_text.replace(f'<a href="">{comment[0]}',f'<span style="text-decoration: underline; color: #56a0d6;"><a href"{comment[1]}" target="_blank" "rel="noopener">{comment[0]}')
                            # print(f"Found match for1: {comment_text}")
                        else:
                            head1_flag=False
                            for i, heading1 in enumerate(heading1_array):
                                if heading1 in comment[1]:
                                    # print("見出し1へのリンクを発見")
                                    blue_text = blue_text.replace(f'<a href="">{comment[0]}',f'<a href="#section{i+1}">{comment[0]}')
                                    head1_flag=True
                                    break
                            if not head1_flag:
                                text_count=0
                                for row, sublist in enumerate(heading2_array):
                                    for col, heading2 in enumerate(sublist):
                                        text_count += 1
                                        if heading2 in comment[1]:
                                            # print("見出し2へのリンクを発見")
                                            blue_text = blue_text.replace(f'<a href="">{comment[0]}',f'<a href="#text{text_count}">{comment[0]}')
                            # print(f"Found match for2: {comment_text}")
                        break
            else:
                blue_text += smart_text
        if prev_is_normal and not is_blue_color(next_word_range) or next_is_table:
            if last_text_count < len(box_last_text) and f'{box_last_text[last_text_count]}' in blue_text:
                blue_text = blue_text.replace('<a href=""></a>','')
                # 新しいリンク番号の付与
                if '<p><a href="">・' in blue_text:
                    blue_text = blue_text.replace('<br />','')
                    list_items = blue_text.split('・')[1:]  # 先頭の"・"を除く
                    blue_text = '<div class="solution" style="padding:10px 15px;border:1px solid #000000;">\r'
                    for idx, item in enumerate(list_items, start=1):
                        list_item = f'<li><a href="#text{h3_count+idx - 1}">{item.strip()}</a></li>'
                        blue_text += list_item + '\r'
                        # h3_count += 1
                    blue_text += '</div>'
                else:
                    blue_text = blue_text.replace('<p>', '<div class="solution" style="padding:10px 15px;border:1px solid #000000;">') + '</a></div>\r'
                    blue_text = blue_text.replace('<br />\r</a></div>','\r</a></div>').replace('<a href="">\r<a href="">','<a href="">').replace('<a href="">\n<a href="">','<a href="">')
                last_text_count += 1
                prev_is_normal = False
            else:
                smart_text.replace('\r', '\n')
                for comment in comments_list:
                    comment_text = comment[0]  # 「コメント対象のテキスト」を取得
                    if f"{comment_text}" in blue_text:
                        # 条件を満たした場合の処理をここに記述
                        if "http" in comment[1]:
                            if "https://www.nurse-step.com/" in comment[1]:
                                link = comment[1]
                                blue_text = blue_text.replace(f'<a href="">{comment[0]}',f'<a href="{link}">{comment[0]}')
                            else:
                                blue_text = blue_text.replace(f'<a href="">{comment[0]}',f'<a href="{comment[1]}" target="_blank">{comment[0]}')
                            # print(f"Found match for3: {comment_text}")
                        else:
                            head1_flag=False
                            for i, heading1 in enumerate(heading1_array):
                                # print(f"heading1:{heading1},comment[1]:{comment[1]}")
                                comment_link_1=comment[1]
                                cleaned_comment_link_1 = remove_surrounding_text(comment_link_1)
                                if is_similar(heading1, cleaned_comment_link_1):
                                    # print("見出し1へのリンクを発見")
                                    blue_text = blue_text.replace(f'<a href="">{comment[0]}',f'<a href="#heading{i+1}">{comment[0]}')
                                    head1_flag=True
                                    break
                            if not head1_flag:
                                for row, sublist in enumerate(heading2_array):
                                    for col, heading2 in enumerate(sublist):
                                        # print(f"heading2:{heading2},comment[1]:{comment[1]}")
                                        if heading2 in comment[1]:
                                            # print("見出し2へのリンクを発見")
                                            blue_text = blue_text.replace(f'<a href="">{comment[0]}',f'<a href="#heading{row+1}_{col+1}">{comment[0]}')
                                            break
                            # print(f"Found match for4: {comment_text}")
                        break
                if is_end(next_word_range):
                    if word_range.Bold:
                        blue_text += f'</a></strong></p>\r'
                    else:
                        blue_text += f'</a></p>'
                    prev_is_normal = False
                else:
                    if word_range.Bold and next_word_range is not None and not next_word_range.Bold:
                        blue_text += f'</a></strong>'
                    else:
                        blue_text += f'</a>'
            blue_text = blue_text.replace('<a href=""></a>','').replace('</a>\r</li>','</a></li>')
            paragraph_text += blue_text
            blue_text = ''
            normal_text = ''
    # ハイライトの開始
    elif is_yellow_color(word_range):
        # bold_text=remove_duplicate_numbers_with_ret(bold_text)
        highlighted_text += normal_text
        normal_text = ''
        if (prev_word_range is not None and not is_end(prev_word_range) and prev_word_range.Bold) and not is_yellow_color(prev_word_range):
            bold_text += '</strong>'
            # paragraph_text += bold_text
            highlighted_text += bold_text
            bold_text = ''
        if prev_word_range is not None and (is_end(prev_word_range) or (not is_yellow_color(prev_word_range)) or is_heading1(prev_word_range) or is_heading2(prev_word_range)):#前回が太字でない場合
            smart_text.replace('\r', '\n')
            if not prev_is_normal:
                highlighted_text += f'<p> <span style="background:linear-gradient(transparent 60%, #ffe57f 60%);">{smart_text}'
                prev_is_normal = True
            else:
                if prev_word_range.Bold:
                    bold_text = ''
                highlighted_text += f' <span style="background:linear-gradient(transparent 60%, #ffe57f 60%);">{smart_text}'
        else:
            highlighted_text += smart_text

        if prev_is_normal and ((next_word_range is not None and not is_yellow_color(next_word_range)) and not is_heading1(next_word_range) and not is_heading2(next_word_range) or is_end(next_word_range)):
            highlighted_text = highlighted_text.replace(' <span style="background:linear-gradient(transparent 60%, #ffe57f 60%);"></span>','').replace('</span> <span style="background:linear-gradient(transparent 60%, #ffe57f 60%);">','')
            # print(highlighted_text)
            if last_text_count < len(box_last_text) and (is_similar(f'{box_last_text[last_text_count]}' , highlighted_text.replace('</span>','').replace(' <span style="background:linear-gradient(transparent 60%, #ffe57f 60%);">','').replace('<strong>','').replace('</strong>','')) or f'{box_last_text[last_text_count]}' in highlighted_text.replace('</span>','').replace(' <span style="background:linear-gradient(transparent 60%, #ffe57f 60%);">','').replace('<strong>','').replace('</strong>','')):#{smart_text}がない
                highlighted_text = highlighted_text.replace('<p>', '<div class="solution" style="padding:10px 15px;border:1px solid #000000;">') + '</span></div>\r'
                highlighted_text = highlighted_text.replace('<br />\r</span></div>','\r</span></div>')
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
                    if last_text_count < len(box_text) and box_count < len(box_text[last_text_count]) and f'{box_text[last_text_count][box_count]}' in highlighted_text.replace('</span>','').replace(' <span style="background:linear-gradient(transparent 60%, #ffe57f 60%);">','').replace('<strong>','').replace('</strong>',''):
                        highlighted_text += f'</span><br />\r'
                        box_count += 1
                    else:
                        highlighted_text += f'</span></p>\r'
                        prev_is_normal = False
                        paragraph_text += normal_text + highlighted_text
                        highlighted_text = ''
                        normal_text = ''
                        box_count = 0
                else:
                    highlighted_text += f'</span>'#ここで<p></strong>が発生している
                    paragraph_text += normal_text + highlighted_text
                    highlighted_text = ''
                    normal_text = ''
                    box_count = 0
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
                    bold_text = bold_text.replace('<p>', '<div class="solution" style="padding:10px 15px;border:1px solid #000000;">') + '</strong></div>\r'
                    bold_text = bold_text.replace('<br />\r</strong></div>','\r</strong></div>')
                    bold_text = bold_text.replace('\n','\r').replace('<br />','<br />\r').replace('\r\r','\r')
                    last_text_count += 1
                    box_count = 0
                    prev_is_normal = False
                    # print(bold_text)
                    paragraph_text += normal_text + bold_text
                    bold_text = ''
                    normal_text = ''
                else:
                    if is_end(next_word_range) or is_heading1(next_word_range) or is_heading2(next_word_range) or'。'in f"{smart_text}":
                        if last_text_count < len(box_text) and box_count < len(box_text[last_text_count]) and f'{box_text[last_text_count][box_count]}' in bold_text:
                            bold_text += f'</strong><br />\r'
                            box_count += 1
                        else:
                            bold_text += f'</strong></p>\r'
                            prev_is_normal = False
                            paragraph_text += normal_text + bold_text
                            bold_text = ''
                            normal_text = ''
                            box_count = 0
                    else:
                        bold_text += f'</strong>'#ここで<p></strong>が発生している
                        paragraph_text += normal_text + bold_text
                        bold_text = ''
                        normal_text = ''
                        box_count = 0
    # マーカーや青色のテキスト以外のテキスト
    elif next_word_range is not None and is_end(next_word_range):
        normal_text += f'{smart_text}'
        if not prev_is_normal:
            paragraph_text += f'<p>{smart_text}</p>'
        elif last_text_count < len(box_last_text) and f'{box_last_text[last_text_count]}' in normal_text:#{smart_text}がない
            normal_text = normal_text.replace('\n','\r')
            normal_text = process_normal_text(normal_text)
            last_text_count += 1
            box_count = 0
            paragraph_text += normal_text
            normal_text = ''
            prev_is_normal = False
            # print(normal_text)
        elif last_text_count < len(box_text) and box_count < len(box_text[last_text_count]) and f'{box_text[last_text_count][box_count]}' in normal_text:
            normal_text += f'<br />'
            box_count += 1
        else:
            # elif last_text_count < len(box_text) and box_count < len(box_text[last_text_count]) and '。' in f'{box_text[last_text_count][box_count]}':
            #     # 「。」でテキストを分割する
            #     split_texts = box_text[last_text_count][box_count].split('。')
            #     match_found = False
            #     combined_text = ""

            #     for i, split_text in enumerate(split_texts):
            #         combined_text += split_text + '。'
            #         if combined_text in normal_text:  # some_target_textは比較対象のテキスト
            #             match_found = True
            #             break
            #     if match_found:
            #         # 一致した場合、</p>を閉じない
            #         pass  # ここで必要な処理を行う
            #     else:
            #         normal_text += f'</p>\r'
            #         paragraph_text += normal_text
            #         normal_text = ''
            #         prev_is_normal = False
            #         box_count = 0
            if "前職では内科のクリニックに7年間勤め、コミュニケーション能力を養ってきました。患者さまとは短い時間の交流ですが、積極的な声かけを継続して関係性を築き、多くの患者さまから声をかけていただけるようになりました。" in normal_text:
                print(normal_text)
                print(f"last_text_count:{last_text_count}")
                print(f"box_count:{box_count}")
                print(f"box_text[last_text_count][box_count]{box_text[last_text_count][box_count]}")
            normal_text += f'</p>\r'
            paragraph_text += normal_text
            normal_text = ''
            prev_is_normal = False
            box_count = 0
    elif prev_is_normal: #次で終わりでないが、テキストはすでに始まっている場合
        if f'<p>▼関連記事はこちら' in normal_text:
            normal_text=f'<p>▼関連記事はこちら<br />\r'
        else:
            normal_text += f'{smart_text}'
            if last_text_count < len(box_last_text) and f'{box_last_text[last_text_count]}' in normal_text:#{smart_text}がない
                normal_text = normal_text.replace('\n','\r')
                normal_text = process_normal_text(normal_text)
                last_text_count += 1
                box_count = 0
                paragraph_text += normal_text
                normal_text = ''
                prev_is_normal = False
                # print(normal_text)
        # prev_is_normal = True
    else:#テキストが始まっていない場合
        if not is_end(word_range):
            normal_text += f'<p>{smart_text}'
            prev_is_normal = True
    # else:
    #     paragraph_text += f'{smart_text}'#なぜか数十個の\rが表示される,本来is_end()で引っかかるはず
    return paragraph_text, normal_text, bold_text, blue_text, highlighted_text,hyper_link_text, hyper_text_count, box_text, box_count, box_last_text, last_text_count, comments_list, links_count, h2_count,prev_is_normal

def extract_text_with_markup(docx_file, html_tables,border_file_path,hyper_links_file_path,links_file_path,heading1_file_path,heading2_file_path):
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
    # docx_file_path_2 を生成
    docx_raw_file_name = os.path.basename(docx_raw_file_path)
    # 生成されるdocxファイルのパスを変更 (例: outputフォルダに保存)
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


    print(f"Trying to open: {docx_file_path_2}")

    # テスト用にファイルの存在確認
    if not os.path.exists(docx_file_path_2):
        raise FileNotFoundError(f"File not found: {docx_file_path_2}")
    
    # テキストファイルの内容を読み込み、リストに保存
    if not os.path.exists(border_file_path):
        raise FileNotFoundError(f"File not found: {border_file_path}")
    
    box_text = border_text.txt_to_2d_array(border_file_path)
    # print(f"box_text[0][0]:{box_text[0][0]}")
    box_count = 0

    #箱の最後のテキストを取り出す
    box_last_text = border_last_text.process_text_file(border_file_path)
    # print(f"{box_last_text}")
    last_text_count = 0

    if not os.path.exists(hyper_links_file_path):
        raise FileNotFoundError(f"File not found: {hyper_links_file_path}")
    hyper_link_text = get_hyper_link_text.read_text_file_to_list(hyper_links_file_path)
    #print(hyper_link_text[0])
    hyper_text_count = 0

    comments_list = []
    comments_list = load_comments_from_file(links_file_path)
    links_count = 0
    # print(f"comments_list[0][0]:{comments_list[0][0]}")

    heading1_array = []
    heading2_array = []
    heading1_array, heading2_array = load_headings(heading1_file_path, heading2_file_path)
    # print(f"heading1_array[0]:{heading1_array[0]}")
    # print(f"heading2_array[0]:{heading1_array[0]}")


    # print(f"box_last_text[0]:{box_last_text[0]}")

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
    h2_count = 0
    h3_count = 1

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
                smart_text = f'{word_range.Text}'
                if word_is_duplicate:
                    smart_text = remove_digits_from_word_range(word_range)
                if prev_word_range is not None:
                    smart_text = remove_duplicate_english_words(word_range,prev_word_range,smart_text)
                if is_heading1(word_range):
                    normal_text = ''
                    bold_text = ''
                    if not is_heading1(prev_word_range):
                        h2_count += 1
                        paragraph_text += f'<h2 id="section{h2_count}">'
                    # 見出し1スタイルのテキストである場合の処理
                    h3_text += smart_text
                    if next_word_range is not None and not is_heading1(next_word_range) or next_is_table:
                        h3_text += f"</h2>\r"
                        paragraph_text += h3_text
                        h3_text = ''
                        paragraph_text, normal_text, bold_text, blue_text, highlighted_text,hyper_link_text, hyper_text_count, box_text, box_count, box_last_text, last_text_count, comments_list, links_count, h2_count, prev_is_normal = check_tag(
            prev_word_range, word_range, next_word_range, paragraph_text, normal_text, bold_text, blue_text, highlighted_text,hyper_link_text, hyper_text_count, box_text, box_count, box_last_text,last_text_count, comments_list, links_count,heading1_array,heading2_array, h2_count,h3_count, prev_is_normal, next_is_table, smart_text)
                elif is_heading2(word_range):
                    normal_text = ''
                    bold_text = ''
                    if not is_heading2(prev_word_range):
                        paragraph_text += f'<h3 id="text{h3_count}">'
                        h3_count += 1
                    # 見出し1スタイルのテキストである場合の処理
                    h4_text += smart_text
                    if next_word_range is not None and not is_heading2(next_word_range) or next_is_table:
                        h4_text += f"</h3>\r"
                        paragraph_text += h4_text
                        h4_text = ''
                        paragraph_text, normal_text, bold_text, blue_text, highlighted_text,hyper_link_text, hyper_text_count, box_text, box_count, box_last_text,last_text_count, comments_list, links_count, h2_count, prev_is_normal = check_tag(
            prev_word_range, word_range, next_word_range, paragraph_text, normal_text, bold_text, blue_text, highlighted_text,hyper_link_text, hyper_text_count, box_text, box_count, box_last_text, last_text_count, comments_list, links_count, heading1_array,heading2_array, h2_count,h3_count, prev_is_normal, next_is_table, smart_text)
                else:
                    paragraph_text, normal_text, bold_text, blue_text, highlighted_text,hyper_link_text, hyper_text_count, box_text, box_count, box_last_text,last_text_count, comments_list, links_count, h2_count, prev_is_normal = check_tag(
            prev_word_range, word_range, next_word_range, paragraph_text, normal_text, bold_text, blue_text, highlighted_text,hyper_link_text, hyper_text_count, box_text, box_count, box_last_text, last_text_count, comments_list, links_count, heading1_array,heading2_array, h2_count,h3_count, prev_is_normal, next_is_table, smart_text)
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
            paragraph_text = paragraph_text.replace('<p>/</p>','').replace('<strong>\r</strong>','').replace('<strong></strong>','').replace('</strong><strong>','').replace('<p>\r</p>','').replace('<p>\r','<p>').replace('','').replace('','').replace('','').replace('<p>▼関連記事はこちら</p>\r<p>','<p>▼関連記事はこちら<br />\r').replace('<p>/</p>','').replace('<br />\r</span></div>','\r</span></div>').replace('<p> </p>','').replace('</a></p>\r<p><a','</a><br />\r<a').replace('</a></p>\n<p><a',',</a><br />\r<a').replace('</a></p>\n<p>',',</a><br />\r').replace('▼関連記事はコチラ</p>\r<p><a','▼関連記事はコチラ<br />\r<a').replace('▼関連記事はコチラ</p>\n<p><a>','▼関連記事はコチラ<br />\n<a').replace('<a href="">\r</a>','').replace('\r</a>','</a>').replace('\n</a>','</a>').replace('<a href="">\r<a href="">','<a href="">').replace('\r</h2>','</h2>').replace('\r</h3>','</h3>').replace('<div class="solution" style="padding:10px 15px;border:1px solid #000000;">\r<br />','<div class="solution" style="padding:10px 15px;border:1px solid #000000;">\r').replace('<p></p>','')
            paragraph_text += '<div class="cta-entry"><div class="cta-enty-inner"><div class="cta-entry_txt"><p>ナースステップは<br><b>「なぜ転職するのか」</b>の理由を<br>明確にするところから一緒に考えます。</p><p><span>看護師の転職に悩んだら<br>まずはナースステップにご相談ください！</span></p></div><div class="cta-entry_btn"><a href="/entry/"><span>無料</span>60 秒で完了！<br>ナースステップに登録する▶</a></div></div></div>'
            while (hyper_text_count < len(hyper_link_text)):
                for comment in comments_list:
                    if hyper_link_text[hyper_text_count] in comment[0]:
                        paragraph_text = paragraph_text.replace(f'{hyper_link_text[hyper_text_count]}',f'<a href="{comment[1]}" target="_blank">{hyper_link_text[hyper_text_count]}</a>')
                        hyper_text_count += 1
                        print(4)
                        break
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


# コマンドライン引数をパースするための設定
parser = argparse.ArgumentParser(description='Process some files.')
parser.add_argument('--config', required=True, help='Path to the config JSON fhle')
args = parser.parse_args()

# JSONファイルのパスを取得
json_file_path = args.config
with open(json_file_path, 'r', encoding='utf-8-sig') as file:
    data = json.load(file)

base_dir = os.path.dirname(json_file_path)
# 元のdocxファイルのパスを取得
docx_raw_file_path = os.path.abspath(os.path.join(base_dir, data["docx_raw_file_path"]))
table_file_path = os.path.abspath(os.path.join(base_dir, data["table_file_path"]))
output_file_path = os.path.abspath(os.path.join(base_dir, data["output_file_path"]))
border_file_path = os.path.abspath(os.path.join(base_dir, data["border_file_path"]))
hyper_links_file_path = os.path.abspath(os.path.join(base_dir, data["hyper_links_file_path"]))
links_file_path = os.path.abspath(os.path.join(base_dir, data["links_file_path"]))
heading1_file_path = os.path.abspath(os.path.join(base_dir, data["heading1_file_path"]))
heading2_file_path = os.path.abspath(os.path.join(base_dir, data["heading2_file_path"]))

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



# HTMLファイルからテーブルを読み込み、リストとして管理
html_tables = read_html_tables(table_file_path)

# WordファイルのテキストとHTMLテーブルを統合して抽出
extracted_text_with_markup = extract_text_with_markup(docx_file_path_2, html_tables,border_file_path,hyper_links_file_path,links_file_path,heading1_file_path,heading2_file_path)

print(99)
html_output = ''.join(extracted_text_with_markup)
with open(output_file_path, 'w', encoding='utf-8') as html_file:
    html_file.write(html_output)
print(100)

print(f'HTMLファイル "{output_file_path}" に抽出されたテキストが保存されました。')