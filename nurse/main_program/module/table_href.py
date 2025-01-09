import re

def replace_href(html):
    # <a>タグごとに処理
    tags = re.split(r'(<a.*?>.*?</a>)', html, flags=re.DOTALL)
    for i, tag in enumerate(tags):
        if 'href=""' in tag:
            # マッチした<a>タグの中のhref=""を置換
            tags[i] = re.sub(r'href=""', lambda match: f'href="#text{replace_href.counter}"', tag)
            replace_href.counter += 1
    return ''.join(tags)

def replace_counter(html, h3_count): 
    replace_href.counter = h3_count
    # 全体を<a>タグごとに分割しつつ処理
    if 'href=""' in html:
        html = replace_href(html)  # 一度に全体を置換
    return html

# 例: HTMLファイルを読み込む部分
def read_html_tables(html_file_path):
    with open(html_file_path, 'r', encoding='utf-8') as file:
        content = file.read()
    tables = re.findall(r'<table.*?</table>', content, re.DOTALL)
    return tables

# 使用例
table_index = 0
h3_count = 1
html_tables = read_html_tables(r'C:\Users\takum\OneDrive\ドキュメント\AWA\nurse\output\4\combined_tables_4.html')

if table_index < len(html_tables):
    html_tables[table_index] = replace_counter(html_tables[table_index], h3_count)
    print(html_tables[table_index])
