import os
import win32com.client as win32

# Word の定数を定義
wdBorderTop = 1
wdBorderBottom = 3
wdBorderLeft = 4
wdBorderRight = 2
wdWithInTable = 12

def remove_table_of_contents(file_path):
    # Word アプリケーションの起動
    word = win32.Dispatch("Word.Application")
    word.Visible = False  # Wordを非表示で実行

    # ドキュメントを編集可能モードで開く
    document = word.Documents.Open(file_path)

    try:
        # 目次をすべて削除
        for toc in document.TablesOfContents:
            toc.Delete()

        # 変更を保存
        output_file_path = file_path.replace('.docx', '_without_toc.docx')
        document.SaveAs(output_file_path)
        print(f"変更が保存されました: {output_file_path}")
    except Exception as e:
        print(f"エラーが発生しました: {e}")
    finally:
        document.Close(False)
        word.Quit()

# 使用例
script_directory = os.path.dirname(os.path.abspath(__file__))
input_file_path = os.path.abspath(os.path.join(script_directory, '..','..', 'input', '240418【校了】自己分析_看護師_転職成功.docx'))
remove_table_of_contents(input_file_path)
