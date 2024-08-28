def read_text_file_to_list(file_path):
    # テキストファイルから一行ずつ読み込み、配列に格納
    lines = []
    with open(file_path, 'r', encoding='utf-8') as file:
        for line in file:
            # 各行の前後の空白や改行を削除してリストに追加
            lines.append(line.strip())
    return lines

# 使用例
# output_file_path = 'hyperlinks_text_output_08_1.txt'  # 作成したテキストファイルのパス
# lines_list = read_text_file_to_list(output_file_path)

# # リスト内容の確認
# for i, line in enumerate(lines_list):
#     print(f"Line {i + 1}: {line}")
