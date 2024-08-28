# AWA

## Automatic Writing Application

wordからhtmlを生成

## フォルダ構成
/input（wordファイルをここに入れてください、githubにはあげていません）

/main_program

┗/module （タイトル・箱・表情報取得のプログラムが入ったフォルダ）

┗awa.py（メインプログラム）

/shell_script（shellコマンドでないと動かせないプログラム）

run_scripts_with_pause.ps1（pythonを一括で動かすshellスクリプト）

README.md（説明ファイル、コマンドが書いてある）

## 実行コマンド
Windows Powershellで以下を実行

python .\main_program\module\get_hyper_links.py

python .\main_program\module\border-checker.py

python .\main_program\module\create_tables.py

python .\main_program\awa.py
