import os
import win32com.client as win32

# Word の定数を定義
wdBorderTop = 1
wdBorderBottom = 3
wdBorderLeft = 4
wdBorderRight = 2
wdWithInTable = 12

def get_text_with_borders():
    try:
        # スクリプトのディレクトリを基準にした相対パスを取得
        script_directory = os.path.dirname(os.path.abspath(__file__))
        docx_file_path = os.path.abspath(os.path.join(script_directory, '..', '..', 'input', '240527_1.docx'))

        # 入力ファイルが存在するか確認
        if not os.path.exists(docx_file_path):
            print(f"ファイルが見つかりません: {docx_file_path}")
            return None

        # Word アプリケーションを作成
        word = win32.Dispatch("Word.Application")
        word.Visible = False

        # Word ドキュメントを開く
        document = word.Documents.Open(docx_file_path)

                # 境界線付き段落のテキストを保存するためのリストを作成
        bordered_texts = []
        current_bordered_text = ""

        # ドキュメント内のすべての段落を反復処理
        paragraph_index = 1  # 段落インデックス
        for para in document.Paragraphs:
            borders = para.Range.Borders

            # 表に含まれている段落をスキップ
            if para.Range.Information(wdWithInTable):
                print(f"スキップされました (表の段落 {paragraph_index}): '{para.Range.Text.strip()}'")
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
            print(f"段落 {paragraph_index} 内容: '{clean_text}'")

            # 段落が "目次" の場合はスキップ
            if clean_text.startswith("目次"):
                print(f"スキップされました (段落 {paragraph_index}): '{clean_text}'")
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


# 関数の使用例
bordered_text_content = get_text_with_borders()

# スクリプトのディレクトリを基準にした相対パスを取得
script_directory = os.path.dirname(os.path.abspath(__file__))
output_file_path = os.path.abspath(os.path.join(script_directory, '..','..', 'output', 'get_border_text_05_1.html'))

# 出力ディレクトリが存在しない場合は作成
output_directory = os.path.dirname(output_file_path)
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

# ファイルに出力
if bordered_text_content:
    with open(output_file_path, 'w', encoding='utf-8') as f:
        f.write(bordered_text_content)

print(f"罫線付き段落のテキストがファイルに保存されました: {output_file_path}")
