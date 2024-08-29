from lxml import etree
import zipfile
import os

def extract_hyperlink_texts(docx_path):
    # Word文書の中のリレーションファイル（XML）を抽出
    with zipfile.ZipFile(docx_path, 'r') as docx:
        # Word文書の内容を取得
        xml_content = docx.read('word/document.xml')
        tree = etree.fromstring(xml_content)

    # ハイパーリンクのテキストを格納するセット（重複排除）
    hyperlink_texts = set()

    # ハイパーリンクの要素を検索
    for elem in tree.iter():
        if elem.tag.endswith('hyperlink'):
            # ハイパーリンクのテキストを取得
            text = ''.join([node.text for node in elem.iter() if node.text])
            if text:
                hyperlink_texts.add(text)  # 重複を防ぐためセットに追加

    return hyperlink_texts

def save_hyperlink_texts_to_file(hyperlink_texts, output_file_path):
    # ファイルに出力（ハイパーリンクのテキストがない場合でもファイルを作成）
    with open(output_file_path, 'w', encoding='utf-8') as file:
        if hyperlink_texts:
            for text in hyperlink_texts:
                file.write(f"{text}\n")
            print(f"罫線付き段落のテキストがファイルに保存されました: {output_file_path}")
        else:
            file.write("")  # 空のファイルを作成

# 使用例
script_directory = os.path.dirname(os.path.abspath(__file__))
input_file_path = os.path.abspath(os.path.join(script_directory, '..', '..', 'input', '240418【校了】自己分析_看護師_転職成功_without_toc_final_no_images.docx'))
output_file_path = os.path.abspath(os.path.join(script_directory, '..', '..', 'output', 'hyperlinks_text_output_04_2.txt'))

hyperlink_texts = extract_hyperlink_texts(input_file_path)
save_hyperlink_texts_to_file(hyperlink_texts, output_file_path)
