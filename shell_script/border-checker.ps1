function Get-TextWithBorders {
    # コンソール出力のエンコードをUTF-8に設定
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8

    try {
        # Word アプリケーションを作成
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false

        # スクリプトのディレクトリを基準にした相対パスを取得
        $scriptDirectory = $PSScriptRoot
        $documentPath = Join-Path -Path $scriptDirectory -ChildPath "../input/240725_3.docx"

        # 入力ファイルが存在するか確認
        if (-not (Test-Path $documentPath)) {
            Write-Host "ファイルが見つかりません: $documentPath"
            return $null
        }

        # Word ドキュメントを開く
        $document = $word.Documents.Open($documentPath)

        # 境界線付き段落のテキストを保存するための配列を作成
        $borderedTexts = @()
        $currentBorderedText = ""

        # ドキュメント内のすべての段落を反復処理
        $paragraphIndex = 1 # 段落インデックス
        foreach ($para in $document.Paragraphs) {
            $borders = $para.Range.Borders

            # 段落に外枠の罫線があるかどうかをチェック
            $hasOutsideBorder = $borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderTop).LineStyle -ne [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone -and
                                $borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderBottom).LineStyle -ne [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone -and
                                $borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderLeft).LineStyle -ne [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone -and
                                $borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderRight).LineStyle -ne [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone

            # 段落がテーブルに含まれていないかどうかをチェック
            $isInTable = $para.Range.Information([Microsoft.Office.Interop.Word.WdInformation]::wdWithInTable)

            # テキストから特定の文字を削除し、トリム
            $cleanText = $para.Range.Text -replace "", "" -replace "\r", "" -replace "\n", "" | ForEach-Object { $_.Trim() }

            # デバッグ用: 各段落のインデックスとテキストを出力
            Write-Host "段落 $paragraphIndex 内容: '$cleanText'"

            # 段落が "目次" の場合はスキップ
            if ($cleanText -match "^目次.*") {
                Write-Host "スキップされました (段落 $paragraphIndex): '$cleanText'"
                $paragraphIndex++
                continue
            }

            # 罫線のある段落を処理
            if ($hasOutsideBorder -and -not $isInTable) {
                if ($currentBorderedText -ne "") {
                    $currentBorderedText += "`n"
                }
                $currentBorderedText += $cleanText
            } else {
                if ($currentBorderedText -ne "") {
                    $borderedTexts += $currentBorderedText
                    $currentBorderedText = ""
                }
            }

            # 段落インデックスを更新
            $paragraphIndex++
        }

        # 最後の罫線付きテキストを追加
        if ($currentBorderedText -ne "") {
            $borderedTexts += $currentBorderedText
        }

        # ドキュメントを閉じる
        $document.Close($false)
        $word.Quit()

        # COM オブジェクトを解放
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($document) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null

        # 境界線付き段落のテキストを連結して返す（段落ごとに改行を追加）
        return $borderedTexts -join "`n`n" # ここで2つの改行を追加

    } catch {
        Write-Error "エラーが発生しました: $_"
    } finally {
        # Wordアプリケーションが開いたままになっている場合は終了
        if ($word -ne $null) {
            $word.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        }
    }
}

# 関数の使用例
$borderedTextContent = Get-TextWithBorders

# スクリプトのディレクトリを基準にした相対パスを取得
$scriptDirectory = $PSScriptRoot
$outputFilePath = Join-Path -Path $scriptDirectory -ChildPath "../output/borderedTextOutput.txt"

# 出力ディレクトリが存在しない場合は作成
$outputDirectory = Split-Path -Parent $outputFilePath
if (-not (Test-Path $outputDirectory)) {
    New-Item -Path $outputDirectory -ItemType Directory | Out-Null
}

# ファイルに出力
$borderedTextContent | Set-Content -Path $outputFilePath -Encoding UTF8

Write-Host "罫線付き段落のテキストがファイルに保存されました: $outputFilePath"
