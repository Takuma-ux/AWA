import os
import win32com.client as win32
import json
import argparse

# Word の定数を定義
wdBorderTop = 1
wdBorderBottom = 3
wdBorderLeft = 4
wdBorderRight = 2
wdWithInTable = 12

def get_text_with_borders(docx_file_path_2):
    try:

        # 入力ファイルが存在するか確認
        if not os.path.exists(docx_file_path_2):
            print(f"ファイルが見つかりません: {docx_file_path_2}")
            return None

        # Word アプリケーションを作成
        word = win32.Dispatch("Word.Application")
        word.Visible = False

        # Word ドキュメントを開く
        document = word.Documents.Open(docx_file_path_2)

                # 境界線付き段落のテキストを保存するためのリストを作成
        bordered_texts = []
        current_bordered_text = ""

        # ドキュメント内のすべての段落を反復処理
        paragraph_index = 1  # 段落インデックス
        for para in document.Paragraphs:
            borders = para.Range.Borders

            # 表に含まれている段落をスキップ
            if para.Range.Information(wdWithInTable):
                # print(f"スキップされました (表の段落 {paragraph_index}): '{para.Range.Text.strip()}'")
                paragraph_index += 1
                continue

            # 境界線の存在をチェック
            has_outside_border = borders(wdBorderTop).LineStyle != 0 and \
                                 borders(wdBorderBottom).LineStyle != 0 and \
                                 borders(wdBorderLeft).LineStyle != 0 and \
                                 borders(wdBorderRight).LineStyle != 0

            # テキストをトリム
            clean_text = para.Range.Text.strip()

            # デバッグ用: 各段落のインデックスとテキストを出力
            # print(f"段落 {paragraph_index} 内容: '{clean_text}'")

            # 段落が "目次" の場合はスキップ
            if clean_text.startswith("目次"):
                # print(f"スキップされました (段落 {paragraph_index}): '{clean_text}'")
                paragraph_index += 1
                continue

            # 罫線のある段落を処理
            if has_outside_border:
                if current_bordered_text != "":
                    current_bordered_text += "\n"
                current_bordered_text += clean_text
            else:
                if current_bordered_text != "":
                    bordered_texts.append(current_bordered_text)
                    current_bordered_text = ""

            # 段落インデックスを更新
            paragraph_index += 1

        # 最後の罫線付きテキストを追加
        if current_bordered_text != "":
            bordered_texts.append(current_bordered_text)

        # ドキュメントを閉じる
        document.Close(False)
        word.Quit()

        # 境界線付き段落のテキストを連結して返す（段落ごとに改行を追加）
        return "\n\n".join(bordered_texts)

    except Exception as e:
        print(f"エラーが発生しました: {e}")
        return None




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
border_file_path = os.path.abspath(os.path.join(base_dir, data["border_file_path"]))

# docx_file_path_2 を生成
docx_raw_file_name = os.path.basename(docx_raw_file_path)
output_dir = os.path.join(base_dir, "output")
if not os.path.exists(output_dir):
    os.makedirs(output_dir)
docx_file_name_modified = docx_raw_file_name.replace('.docx', '_without_toc_final_no_images.docx')
docx_file_path_2 = os.path.abspath(os.path.join(output_dir, docx_file_name_modified))

# 関数の使用例
bordered_text_content = get_text_with_borders(docx_file_path_2)
with open(border_file_path, 'w', encoding='utf-8') as f:
    # ファイルに出力
    if bordered_text_content:
            f.write(bordered_text_content)
    else:
        f.write("")  # 空のファイルを作成

print(f"罫線付き段落のテキストがファイルに保存されました: {border_file_path}")
