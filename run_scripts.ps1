# コマンドの後に休止時間を設定する関数
function Run-WithPause {
    param (
        [string]$command
    )

    Invoke-Expression $command
    Start-Sleep -Seconds 5  # 5秒の休止時間を設定
}


Run-WithPause "python ./main_program/module/delete_top.py"
Run-WithPause "python ./main_program/module/get_head.py"
Run-WithPause "python ./main_program/module/delete_top_table.py"
Run-WithPause "python ./main_program/module/delete_img.py"
Run-WithPause "python ./main_program/module/get_hyper_links.py"
Run-WithPause "python ./main_program/module/border-checker.py"
Run-WithPause "python ./main_program/module/create_tables.py"
Run-WithPause "python ./main_program/awa_05.py"