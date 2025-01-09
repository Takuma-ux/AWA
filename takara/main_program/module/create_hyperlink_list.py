import json
import argparse
import re
import os
from spire.doc import *
from spire.doc.common import *

# Find all the hyperlinks in a document
def FindAllHyperlinks(document):
    hyperlinks = []
    for i in range(document.Sections.Count):
        section = document.Sections.get_Item(i)
        for j in range(section.Body.ChildObjects.Count):
            sec = section.Body.ChildObjects.get_Item(j)
            if sec.DocumentObjectType == DocumentObjectType.Paragraph:
                for k in range(sec.ChildObjects.Count):
                    para = sec.ChildObjects.get_Item(k)
                    if para.DocumentObjectType == DocumentObjectType.Field:
                        field = para if isinstance(para, Field) else None
                        if field.Type == FieldType.FieldHyperlink:
                            hyperlinks.append((field, sec))
    return hyperlinks


# Convert IEnumerator to list
def ienumerator_to_list(ienumerator):
    return [ienumerator.get_Item(i) for i in range(ienumerator.Count)]

# Get hyperlink text and URL
# Get hyperlink text and URL
def GetHyperlinkDetails(field, paragraph):
    if field.Type == FieldType.FieldHyperlink:
        # ハイパーリンクのテキストを取得（段落から取得）
        text = ""
        if paragraph is not None and hasattr(paragraph, 'ChildObjects'):
            # ChildObjects を安全にリストに変換してからイテレーション
            child_objects = ienumerator_to_list(paragraph.ChildObjects)
            for child in child_objects:
                if isinstance(child, TextRange):
                    text += child.Text

        # ハイパーリンクのURLを取得
        url = ""
        if field.Code:
            url = field.Code.split('"')[1] if '"' in field.Code else ""
        
        return text.strip(), url.strip()
    return "", ""

def RemoveHyperlink(field):
    paragraph = field.OwnerParagraph
    # ハイパーリンクに関連するテキストを取得
    hyperlink_text = ""
    hyperlink_index = -1
    if hasattr(paragraph, 'ChildObjects'):
        child_objects = ienumerator_to_list(paragraph.ChildObjects)
        # ChildObjects内をループし、ハイパーリンクの位置を取得してテキストを保持
        for i, child in enumerate(child_objects):
            if isinstance(child, TextRange):
                hyperlink_text += child.Text
            if child == field:
                hyperlink_index = i

    # ハイパーリンクの Field を削除する前に、TextRange を新しく作成して保持
    new_text_range = TextRange(doc)
    new_text_range.Text = hyperlink_text

    # 段落からハイパーリンクフィールドを削除
    paragraph.ChildObjects.Remove(field)

    # 段落にハイパーリンクテキストを追加（元のハイパーリンクの位置に）
    if hyperlink_index != -1:
        paragraph.ChildObjects.Insert(hyperlink_index, new_text_range)
    else:
        paragraph.ChildObjects.Add(new_text_range)


# Create a Document object
doc = Document()

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
hyperlink_file_path = os.path.abspath(os.path.join(base_dir, data["hyperlink_file_path"]))
# docx_file_path_2 を生成
docx_raw_file_name = os.path.basename(docx_raw_file_path)
# JSONファイル名から数字を抽出
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

doc.LoadFromFile(docx_file_path_2)

# Get all hyperlinks
hyperlinks = FindAllHyperlinks(doc)

# Prepare to save hyperlink texts
hyperlink_data = []
unique_texts = set()  # ハイパーリンクテキストの重複を防ぐためにセットを使用

# # Extract hyperlink details
# for hyperlink, paragraph in hyperlinks:
#     text, url = GetHyperlinkDetails(hyperlink, paragraph)
#     hyperlink_data.append((text, url))

# Extract hyperlink details and remove hyperlinks
for hyperlink, paragraph in hyperlinks:
    text, url = GetHyperlinkDetails(hyperlink, paragraph)
    
    # ハイパーリンクテキストが重複していない場合にのみ追加
    if text not in unique_texts:
        hyperlink_data.append((text, url))
        unique_texts.add(text)
    
    RemoveHyperlink(hyperlink)

# Save hyperlink data to a new text file
with open(hyperlink_file_path, 'w', encoding='utf-8') as f:
    for text, url in hyperlink_data:
        f.write(f'[{text}],[{url}]\n')

# Save the modified document
docx_remove_hyperlinks_path = docx_raw_file_name.replace('.docx', '_without_toc_final_no_images_remove_hyperlinks.docx')
output_docx_file_path = os.path.abspath(os.path.join(output_dir, docx_remove_hyperlinks_path))
doc.SaveToFile(output_docx_file_path, FileFormat.Docx)

doc.Close()