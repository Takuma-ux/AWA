import os
def txt_to_2d_array(file_path):
    array_2d = []
    current_section = []

    with open(file_path, 'r', encoding='utf-8') as file:
        for line in file:
            # line = line.replace(' ', '')  # スペースを削除
            stripped_line = line.strip()
            if stripped_line:  # 空行でない場合
                current_section.append(stripped_line)
            else:
                if current_section:
                    array_2d.append(current_section)
                    current_section = []
        
        # 最後に残っているセクションを追加
        if current_section:
            array_2d.append(current_section)

    return array_2d

# if __name__ == "__main__":
#     # テキストファイルのパス
#     script_directory = os.path.dirname(os.path.abspath(__file__))
#     input_file_path = os.path.abspath(os.path.join(script_directory, '..','..', 'output', 'get_border_text_04_1.html'))
    
#     # 二次元配列としてテキストファイルを処理
#     result_array = txt_to_2d_array(input_file_path)
    
#     # 結果を表示
#     for section in result_array:
#         print(section)
