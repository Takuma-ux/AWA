import os
import win32com.client as win32
import json
import argparse
import re

# 定数を手動で設定
wdYellow = 7
wdNoHighlight = 0

def load_headings(heading1_file):
    heading1_list = []

    # 見出し1のファイルを読み込み
    with open(heading1_file, 'r', encoding='utf-8') as file:
        heading1_list = [line.strip() for line in file if line.strip()]  # 空行を除いてリストに追加

    return heading1_list

def check_editable_range(document, start, end):
    try:
        test_range = document.Range(Start=start, End=end)
        test_range.Text = "test"  # 一時的にテキストを変更してみる
        test_range.Text = ""  # 元に戻す
        return True
    except Exception as e:
        print(f"範囲を編集できません: {e}")
        return False

def remove_before_specific_text_and_insert_heading(file_path, stop_text):
    # Word アプリケーションの起動
    word = win32.Dispatch("Word.Application")
    word.Visible = False  # Wordを非表示で実行

    # ドキュメントを編集可能モードで開く（ReadOnly=Falseで開く）
    document = word.Documents.Open(file_path, ReadOnly=False)

    try:
        # ドキュメントが保護されている場合は解除する
        if document.ProtectionType != -1:
            print("ドキュメントが保護されているため、保護を解除します。")
            document.Unprotect()

        # 目次（TOC）が存在する場合は削除する
        if document.TablesOfContents.Count > 0:
            print("目次を削除します。")
            for toc in document.TablesOfContents:
                toc.Delete()

        range_start = document.Content.Start
        range_end = None

        # 指定されたテキストを検索
        for paragraph in document.Paragraphs:
            if stop_text in paragraph.Range.Text:
                range_end = paragraph.Range.End
                break

        # 範囲のデバッグ情報を表示
        print(f"range_start: {range_start}, range_end: {range_end}")

        if range_end is not None:
            # 範囲が編集可能かどうかをチェック
            if check_editable_range(document, range_start, range_end):
                # 前半部分を削除（`Cut` を使用）
                document.Range(Start=range_start, End=range_end).Cut()

                # stop_text をドキュメントの先頭に挿入し、見出し1スタイルを適用
                new_paragraph = document.Paragraphs.Add(document.Content)
                new_paragraph.Range.Text = f'\n{stop_text}\n'
                new_paragraph.Range.Style = document.Styles("見出し 1")

                # 黄色マーカーのチェックと削除
                if new_paragraph.Range.HighlightColorIndex == wdYellow:
                    new_paragraph.Range.HighlightColorIndex = wdNoHighlight
            else:
                print("指定された範囲は編集できませんでした。")
        else:
            print(f"指定されたテキスト '{stop_text}' が見つかりませんでした。")

        # 変更を保存
        output_file_path = file_path.replace('.docx', '_final.docx')
        document.SaveAs(output_file_path)
        print(f"変更が保存されました: {output_file_path}")
    except Exception as e:
        print(f"エラーが発生しました: {e}")
    finally:
        document.Close(False)
        word.Quit()

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
heading1_file_path = os.path.abspath(os.path.join(base_dir, data["heading1_file_path"]))
# 生成されるdocxファイルのパスを変更 (例: outputフォルダに保存)
# JSONファイル名から数字を抽出
match = re.search(r'\d+', os.path.basename(json_file_path))
if match:
    number = match.group()
else:
    number = 'default'  # 数字が見つからない場合のデフォルト値

output_dir = os.path.join(base_dir, "output", f"{number}")
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

docx_raw_file_name = os.path.basename(docx_raw_file_path)
docx_file_name_modified = docx_raw_file_name.replace('.docx', '_without_toc.docx')
docx_file_path_2 = os.path.abspath(os.path.join(output_dir, docx_file_name_modified))
print(f"docx_file_path_2: {docx_file_path_2}")
heading1_array = []
heading1_array = load_headings(heading1_file_path)

# ファイルが存在するか確認する
if not os.path.exists(docx_file_path_2):
    print(f"エラー: ファイルが存在しません: {docx_file_path_2}")
else:
    print(f"ファイルを処理中: {docx_file_path_2}")
    stop_text = heading1_array[0]  # 例として設定
    print(f"対象のテキスト{stop_text}")
    remove_before_specific_text_and_insert_heading(docx_file_path_2, stop_text)
