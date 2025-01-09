import win32com.client as win32

# Wordファイルを開く関数
def open_word_file(docx_file):
    # Wordアプリケーションを非表示で起動
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    
    # ドキュメントを開く
    doc = word.Documents.Open(docx_file)
    return doc, word

# Wordファイルのテキストを解析する関数
def analyze_word_file(docx_file):
    # Wordファイルを開く
    doc, word = open_word_file(docx_file)
    
    try:
        # Wordの段落とテキストの範囲を取得
        for range in doc.StoryRanges:
            # 各段落内のテキストを結合して1つの文にする
            word_ranges = list(range.Words)  # Wordsをリストに変換
            paragraph_text = ''
            for i, word_range in enumerate(word_ranges):
                next_word_range = None
                prev_word_range = word_ranges[i - 1] if i > 0 else None
                prev_word_range_1 = word_ranges[i - 2] if i > 1 else None
                prev_word_range_2 = word_ranges[i - 3] if i > 2 else None
                prev_word_range_3 = word_ranges[i - 4] if i > 3 else None
                if (i + 1 < len(word_ranges)):
                    next_word_range = word_ranges[i + 1]

                # 出力する場所を特定
                if 'CS' in word_range.Text:
                    print(f"Word Range Text: {repr(word_range.Text)}")
                    print(f"Previous Word Range Text: {repr(prev_word_range.Text if prev_word_range else 'None')}")
                    print(f"Previous Word Range Text: {repr(prev_word_range_1.Text if prev_word_range_1 else 'None')}")
                    print(f"Previous Word Range Text: {repr(prev_word_range_2.Text if prev_word_range_2 else 'None')}")
                    print(f"Previous Word Range Text: {repr(prev_word_range_3.Text if prev_word_range_3 else 'None')}")
                    print(f"Next Word Range Text: {repr(next_word_range.Text if next_word_range else 'None')}")
                    print(f"Current paragraph_text: {repr(paragraph_text)}")
                    break
    
    finally:
        # ドキュメントを閉じ、Wordアプリケーションを終了
        doc.Close()
        word.Quit()

# Wordファイルのパスを指定
docx_file = r'C:\Users\takum\OneDrive\ドキュメント\AWA\takara\output\107\240213【初稿】ビルトインコンロとは_without_toc_final_no_images_remove_hyperlinks.docx'  # Wordファイルのパスを指定してください

# Wordファイルを解析してデバッグ情報を表示
analyze_word_file(docx_file)
