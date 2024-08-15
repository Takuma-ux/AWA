from docx import Document
import os

def remove_images_from_docx(file_path):
    # ドキュメントを読み込む
    try:
        doc = Document(file_path)
    except Exception as e:
        print(f"ファイルを開く際にエラーが発生しました: {e}")
        return
    
    # 各段落内の画像を削除
    for para in doc.paragraphs:
        for run in para.runs:
            if 'Graphic' in run.element.xml:
                run.clear()

    # 各表内の画像を削除
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        if 'Graphic' in run.element.xml:
                            run.clear()

    # 変更を保存
    output_file_path = file_path.replace('.docx', '_no_images.docx')
    try:
        doc.save(output_file_path)
        print(f"ファイルを保存しました: {output_file_path}")
    except Exception as e:
        print(f"ファイルを保存する際にエラーが発生しました: {e}")

    # 保存先のディレクトリとファイル名を確認
    print(f"保存先ディレクトリ: {os.path.dirname(output_file_path)}")
    print(f"保存ファイル名: {os.path.basename(output_file_path)}")

# 使用
file_path = r'C:\Users\nichi\OneDrive\デスクトップ\study\auto\240422【校了】登録販売者_役に立たない.docx'
remove_images_from_docx(file_path)

