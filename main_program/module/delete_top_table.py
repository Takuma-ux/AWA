import os
import win32com.client as win32

# 定数を手動で設定
wdYellow = 7
wdNoHighlight = 0
def load_headings(heading1_file):
    heading1_list = []

    # 見出し1のファイルを読み込み
    with open(heading1_file, 'r', encoding='utf-8') as file:
        heading1_list = [line.strip() for line in file if line.strip()]  # 空行を除いてリストに追加

    return heading1_list
def remove_before_specific_text_and_insert_heading(file_path, stop_text):
    # Word アプリケーションの起動
    word = win32.Dispatch("Word.Application")
    word.Visible = False  # Wordを非表示で実行

    # ドキュメントを編集可能モードで開く
    document = word.Documents.Open(file_path)

    try:
        range_start = document.Content.Start
        range_end = None

        # 指定されたテキストを検索
        for paragraph in document.Paragraphs:
            if stop_text in paragraph.Range.Text:
                range_end = paragraph.Range.End
                break

        if range_end is not None:
            # 前半部分を削除
            document.Range(Start=range_start, End=range_end).Delete()

            # stop_text をドキュメントの先頭に挿入し、見出し1スタイルを適用
            new_paragraph = document.Paragraphs.Add(document.Content)
            new_paragraph.Range.Text = stop_text
            new_paragraph.Range.Style = document.Styles("見出し 1")

            # 黄色マーカーのチェックと削除
            if new_paragraph.Range.HighlightColorIndex == wdYellow:
                new_paragraph.Range.HighlightColorIndex = wdNoHighlight

        # 変更を保存
        output_file_path = file_path.replace('.docx', '_final.docx')
        document.SaveAs(output_file_path)
        print(f"変更が保存されました: {output_file_path}")
    except Exception as e:
        print(f"エラーが発生しました: {e}")
    finally:
        document.Close(False)
        word.Quit()

# 使用例
script_directory = os.path.dirname(os.path.abspath(__file__))
input_file_path = os.path.abspath(os.path.join(script_directory, '..','..', 'input', '240418【校了】自己分析_看護師_転職成功_without_toc.docx'))
heading1_file_path = os.path.abspath(os.path.join(script_directory, '..', '..', 'output', 'heading4_1.txt'))
heading1_array = []
heading1_array = load_headings(heading1_file_path)
print(heading1_array[0])
stop_text = heading1_array[0]
remove_before_specific_text_and_insert_heading(input_file_path, stop_text)
