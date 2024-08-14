# コマンドの後に休止時間を設定する関数
function Run-WithPause {
    param (
        [string]$command
    )

    Invoke-Expression $command
    Start-Sleep -Seconds 5  # 5秒の休止時間を設定
}

# PowerShellスクリプトとPythonスクリプトを順に実行し、休止時間を設ける
Run-WithPause "powershell -File shell_script/border-checker_07.ps1"
Run-WithPause "python ./main_program/awa_re_07.py"

Run-WithPause "powershell -File shell_script/border-checker_06.ps1"
Run-WithPause "python ./main_program/awa_re_06.py"

Run-WithPause "powershell -File shell_script/border-checker_05.ps1"
Run-WithPause "python ./main_program/awa_re_05.py"

Run-WithPause "powershell -File shell_script/border-checker_04.ps1"
Run-WithPause "python ./main_program/awa_re_04.py"

Run-WithPause "powershell -File shell_script/border-checker_03.ps1"
Run-WithPause "python ./main_program/awa_re_03.py"

Run-WithPause "powershell -File shell_script/border-checker_02.ps1"
Run-WithPause "python ./main_program/awa_re_02.py"

Run-WithPause "powershell -File shell_script/border-checker_01.ps1"
Run-WithPause "python ./main_program/awa_re_01.py"
