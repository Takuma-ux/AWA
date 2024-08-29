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

（※Wordファイルを開いていると、エラーが発生します。タスクマネージャーで閉じていることを確認してください）

powershell -ExecutionPolicy Bypass -File ./run_scripts.ps1
