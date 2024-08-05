import os
import win32com.client
import colorsys
from component import get_title
# from component import get_box_last_text_3
from component import border_last_text

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
    # 指定されたWordのRangeが「見出し1」のスタイルかどうかを判定する関数。
    # word_range.StyleがNoneでないことを確認
    if word_range.Style is not None:
        return word_range.Style.NameLocal == "見出し 1"
    return False

def is_heading2(word_range):
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

def extract_text_with_markup(docx_file):
    script_directory = os.path.dirname(os.path.abspath(__file__))
    docx_file_path = os.path.abspath(os.path.join(script_directory, '..', 'input', '240725_2.docx'))
    input_file_path = os.path.abspath(os.path.join(script_directory, '..', 'output', 'borderedTextOutput.txt'))
    print(f"Trying to open: {docx_file_path}")

    # テスト用にファイルの存在確認
    if not os.path.exists(docx_file_path):
        raise FileNotFoundError(f"File not found: {docx_file_path}")
    # Wordファイルからテキストを抽出
    extracted_text = get_title.extract_text_from_docx(docx_file_path)
    # 数字で終わる行を抽出
    lines_ending_with_number = get_title.extract_lines_ending_with_number(extracted_text)
    # 末尾の数字を削除して各行のテキストを取り出す
    cleaned_text = get_title.remove_numbers_from_end(lines_ending_with_number)
    # 末尾の数字を削除して各行のテキストを取り出す
    cleaned_text = [line.rstrip() for line in cleaned_text]
    #箱の最後のテキストを取り出す
    box_last_text = border_last_text.remove_duplicate_numbers_with_ret(input_file_path)
    count = 0
    last_text_count = 0
    sub_title_count = 0
    sub_title_length = 0
    # print("start")
    # # 結果を表示

    for line in cleaned_text:
        print(repr(line))
    # print("end")

    # print(repr(f"<p><strong>{cleaned_text[3]}\r"))

    length=len(cleaned_text)
    print("length:",len(cleaned_text))

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False  # Wordアプリケーションを非表示にする

    doc = word.Documents.Open(docx_file)
    extracted_text = []
    sub_title_list = []

    # 直前の文字列が青色のテキストであるかどうかを示すフラグ
    prev_is_blue = False
    # 直前の文字列が太字のテキストであるかどうかを示すフラグ
    prev_is_bold = False
    # 直前の文字列が黄色のマーカーであるかどうかを示すフラグ
    prev_is_highlighted = False
    # 直前の文字列が普通のテキストであるかどうかを示すフラグ
    prev_is_normal = False
    # 直前の文字列が見出し1であるかどうかを示すフラグ
    prev_is_h3 = False
    # 直前の文字列が見出し2であるかどうかを示すフラグ
    prev_is_h4 = False
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

    for range in doc.StoryRanges:
        # 各段落内のテキストを結合して1つの文にする
        paragraph_text = ''
        for word_range in range.Words:
            if is_heading1(word_range):
                normal_text = ''
                bold_text = ''
                prev_is_bold = False
                prev_is_normal = False
                if not prev_is_h3:
                    paragraph_text += f'<h3>'
                    prev_is_h3 = True
                # 見出し1スタイルのテキストである場合の処理
                h3_text += f"{word_range.Text.strip()}"
            elif prev_is_h3:
                h3_text += f"</h3>\r"
                paragraph_text += h3_text
                h3_text = ''
                prev_is_h3 = False
            if is_heading2(word_range):
                normal_text = ''
                bold_text = ''
                prev_is_bold = False
                prev_is_normal = False
                if not prev_is_h4:
                    paragraph_text += f'<h4>'
                    prev_is_h4 = True
                # 見出し1スタイルのテキストである場合の処理
                h4_text += f"{word_range.Text.strip()}"
            elif prev_is_h4:
                h4_text += f"</h4>\r"
                paragraph_text += h4_text
                h4_text = ''
                prev_is_h4 = False
            # 青色の開始
            if is_blue_color(word_range):
                if not prev_is_blue:
                    if not prev_is_normal:
                        blue_text = f'<p><a href="">{word_range.Text}'
                        prev_is_normal = True
                    else:
                        blue_text = f'<a href="">{word_range.Text}'
                else:
                    blue_text += word_range.Text
                prev_is_blue = True
            # 青色の終了
            elif prev_is_blue:
                blue_text = blue_text.replace('\r','')
                blue_text += f'</a></p>\r'#リンクの場合毎回閉じる
                paragraph_text += blue_text
                blue_text = ''
                normal_text = ''
                prev_is_normal = False
                prev_is_blue = False
            # ハイライトの開始
            elif is_yellow_color(word_range):
                bold_text=remove_duplicate_numbers_with_ret(bold_text)
                paragraph_text += normal_text
                normal_text = ''
                if prev_is_bold:
                    bold_text += '</strong>'
                    paragraph_text += bold_text
                    bold_text = ''
                    prev_is_bold = False
                if not prev_is_highlighted:
                    word_range.Text.replace('\r', '\n')
                    if not prev_is_normal:
                        highlighted_text = f'<p><span class="marker"><strong>{word_range.Text}'
                        prev_is_normal = True
                    else:
                        if prev_is_bold:
                            bold_text = ''
                        highlighted_text = f'<span class="marker"><strong>{word_range.Text}'
                else:
                    highlighted_text += word_range.Text
                prev_is_highlighted = True
            # ハイライトの終了
            elif prev_is_highlighted:
                if f'{box_last_text[last_text_count]}' in highlighted_text:#{word_range.Text}がない
                    sub_title_list = make_list_subtitle(highlighted_text)
                    sub_title_length=len(sub_title_list)
                    print("sub_title_length_highlight",sub_title_length)
                    sub_title_count = 0
                    highlighted_text = highlighted_text.replace('<p>', '').replace('\r','<br />\r').replace('<span class="marker"><strong>', '<div style="background:#ffffff;border:1px solid #cccccc;padding:5px 10px;"><span class="marker"><strong>') + '</strong></span></div>\r'
                    highlighted_text = highlighted_text.replace('<br />\r</strong></span></div>','\r</strong></span></div>')
                    last_text_count += 1
                    # normal_text += f'<p>word_range.Text'
                    prev_is_normal = False
                else:
                    word_range.Text.replace('\r', '\n')
                    if is_end(word_range):
                        highlighted_text += f'</strong></span></p>\r<p>{word_range.Text}'
                        # prev_is_normal = False
                    else:
                        if '。'in f"{word_range.Text}":
                            highlighted_text += f'</strong></span>{word_range.Text}</p>\r<p>'
                        else:
                            highlighted_text += f'</strong></span>{word_range.Text}'
                paragraph_text += highlighted_text
                highlighted_text = ''
                normal_text = ''
                prev_is_highlighted = False
                prev_is_bold = False
            # 太字の開始
            elif word_range.Bold:
                bold_text=remove_duplicate_numbers_with_ret(bold_text)
                if not prev_is_bold:#おかしなところでprev_is_boldが続いてしまっているので、
                    if not prev_is_normal:
                        bold_text = f'<p><strong>{word_range.Text}'
                        prev_is_normal = True
                    else:
                        bold_text = f'<strong>{word_range.Text}'
                    prev_is_bold = True
                else:
                    bold_text += word_range.Text
                    # prev_is_bold = True
                
            # 太字の終了
            elif prev_is_bold:
                bold_text=remove_duplicate_numbers_with_ret(bold_text)
                if is_end(word_range):
                    bold_text += f'</strong></p>\r<p>{word_range.Text}'
                    prev_is_bold = False
                    # prev_is_normal = False
                else:
                    if '。'in f"{word_range.Text}":
                        bold_text += f'</strong>{word_range.Text}</p>\r<p>'
                    else:
                        bold_text += f'</strong>{word_range.Text}'#ここで<p></strong>が発生している
                    paragraph_text += normal_text+bold_text
                    bold_text = ''
                    normal_text = ''
                    prev_is_bold = False
            # マーカーや青色のテキスト以外のテキスト
            elif not is_end(word_range):
                if not prev_is_normal:
                    paragraph_text += f'<p>{word_range.Text}'
                else:
                    if '。' in word_range.Text:
                        normal_text += word_range.Text + '</p>\r<p>' 
                    else:
                        normal_text += word_range.Text #２回目のループからここにずっと来ている
                prev_is_normal = True
            elif prev_is_normal: #ここにも来ない
                if f'<p>▼関連記事はこちら' in normal_text:
                    normal_text=f'<p>▼関連記事はこちら<br />\r'
                else:
                    normal_text += f'</p>{word_range.Text}'
                    paragraph_text += normal_text
                    normal_text = ''
                    prev_is_normal = False
            # else:
            #     paragraph_text += f'{word_range.Text}'#なぜか数十個の\rが表示される,本来is_end()で引っかかるはず

        # テキストが空でない場合のみ処理を行います
        if paragraph_text:
            paragraph_text = paragraph_text.replace('<p>/</p>','').replace('<p></p>','').replace('<strong></strong>','').replace('<p>\r</p>','').replace('<p>\r','<p>').replace('','').replace('','').replace('','').replace('<p>▼関連記事はこちら</p>\r<p>','<p>▼関連記事はこちら<br />\r').replace('<p>/</p>','')
            paragraph_text = paragraph_text.replace('<br />\r</span></div>','\r</span></div>')
            # 文の末尾に数字がある場合、その数字を取り除く
            cleaned_text = remove_trailing_digits(paragraph_text)
            # 連続する数字を1回のみ表示する
            cleaned_text = remove_duplicate_numbers_with_ret(cleaned_text)
            if cleaned_text:
                extracted_text.append(cleaned_text)

    doc.Close()
    word.Quit()
    return extracted_text

script_directory = os.path.dirname(os.path.abspath(__file__))
docx_file_path = os.path.abspath(os.path.join(script_directory, '..', 'input', '240725_2.docx'))

extracted_text_with_markup = extract_text_with_markup(docx_file_path)
html_output = ''.join(extracted_text_with_markup)
output_file_path = os.path.abspath(os.path.join(script_directory, '..', 'output', 'extracted_text_knowledge.html'))
with open(output_file_path, 'w', encoding='utf-8') as html_file:
    html_file.write(html_output)

print(f'HTMLファイル "{output_file_path}" に抽出されたテキストが保存されました。')
