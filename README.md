# AWA

## Automatic Writing Application

wordからhtmlを生成

## フォルダ構成
/input（wordファイルをここに入れてください、githubにはあげていません）

/main_program

┗/module （タイトル・箱・表情報取得のプログラムが入ったフォルダ）

┗awa.py（メインプログラム）

run_scripts.ps1（pythonを一括で動かすshellスクリプト）

README.md（説明ファイル、コマンドが書いてある）

## 注意事項
・Wordファイルを開いていると、エラーが発生します。タスクマネージャーで閉じていることを確認してください

・Wordのスタイル見出し1・2のテキストを設定していないとエラーが発生することがあります

## 実行コマンド
Windows Powershellで以下を実行

powershell -ExecutionPolicy Bypass -File ./run_scripts.ps1
