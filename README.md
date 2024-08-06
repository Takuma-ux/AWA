# AWA

#Automatic Writing Application

#wordからhtmlを生成


/input（wordファイルをここに入れてください、githubにはあげていません）
/main_program
┗/component （タイトル・箱・表情報取得のプログラムが入ったフォルダ）
┗awa.py（過去プログラム）
┗awa_2.py（メインプログラム）
/shell_script（shellコマンドでないと動かせないプログラム）
README.md（説明ファイル、コマンドが書いてある）

Windows Powershellで以下を実行

powershell -File shell_script/border-checker.ps1

python .\main_program\awa_2.py
