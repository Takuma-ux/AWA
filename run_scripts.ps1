# Run-WithPause 関数の定義
function Run-WithPause {
    param (
        [string]$command
    )

    Invoke-Expression $command
    Start-Sleep -Seconds 5  # 5秒の休止時間を設定
}

# inputフォルダをスキャンして、最初の状態ですべての.docxファイルを取得
$inputFolder = "./input"
$docxFiles = Get-ChildItem -Path $inputFolder -Filter "*.docx"

# カウント変数を初期化
$count = 1

# 各.docxファイルについて処理
foreach ($docxFile in $docxFiles) {
    # JSONファイルのパスを設定
    $configJsonPath = "./config_$count.json"

    # JSONの内容を定義 (名前にカウントを追加)
    $jsonContent = @{
        "docx_raw_file_path" = "./input/$($docxFile.Name)"
        "table_file_path" = "./output/combined_tables_$count.html"
        "border_file_path" = "./output/get_border_text_$count.html"
        "hyper_links_file_path" = "./output/hyperlinks_text_output_$count.txt"
        "links_file_path" = "./output/comments_output_$count.txt"
        "heading1_file_path" = "./output/heading1_$count.txt"
        "heading2_file_path" = "./output/heading2_$count.txt"
        "output_file_path" = "./output/extracted_text_$count.html"
    }

    # JSONファイルを作成 (utf-8 エンコードを指定)
    $jsonContent | ConvertTo-Json | Out-File -FilePath $configJsonPath -Encoding utf8

    # 生成されたJSONファイルのパスを表示
    Write-Host "Generated JSON file: $configJsonPath"

    # カウントをインクリメント
    $count++
}
$count = 1
# 生成された各JSONファイルに対してPythonスクリプトを実行
foreach ($docxFile in $docxFiles) {
    $configJsonPath = "./config_$count.json"
    Run-WithPause "python ./main_program/module/delete_top.py --config $configJsonPath"
    Run-WithPause "python ./main_program/module/get_head.py --config $configJsonPath"
    Run-WithPause "python ./main_program/module/delete_top_table.py --config $configJsonPath"
    Run-WithPause "python ./main_program/module/delete_img.py --config $configJsonPath"
    Run-WithPause "python ./main_program/module/get_hyper_links.py --config $configJsonPath"
    Run-WithPause "python ./main_program/module/border-checker.py --config $configJsonPath"
    Run-WithPause "python ./main_program/module/get_links.py --config $configJsonPath"
    Run-WithPause "python ./main_program/module/create_tables.py --config $configJsonPath"
    Run-WithPause "python ./main_program/awa.py --config $configJsonPath"
    $count++
}
