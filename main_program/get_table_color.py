from docx import Document
import xml.etree.ElementTree as ET
import os

def get_cell_background_color(cell):
    # cell.xml はセルの XML 表現
    cell_xml = cell._tc
    # XML ツリーを解析
    tree = ET.ElementTree(ET.fromstring(cell_xml.xml))
    # セルのシェーディングを検索
    shd_elem = tree.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd')
    if shd_elem is not None:
        # 塗りつぶしの色を取得
        fill_color = shd_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill')
        return fill_color
    return None

def print_tables_with_colors(docx_path):
    doc = Document(docx_path)
    for table_index, table in enumerate(doc.tables):
        print(f"Table {table_index + 1}:")
        for row_index, row in enumerate(table.rows):
            for cell_index, cell in enumerate(row.cells):
                text = cell.text
                # セルの背景色を取得
                background_color = get_cell_background_color(cell)
                if background_color:
                    print(f"Row {row_index + 1}, Cell {cell_index + 1}: '{text}' with background color {background_color}")
                else:
                    print(f"Row {row_index + 1}, Cell {cell_index + 1}: '{text}' with no background color")

# 相対パスを設定
relative_path = os.path.join("input", "test_0802.docx")
print_tables_with_colors(relative_path)
