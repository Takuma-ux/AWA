import spire.doc
import os

# バージョンを取得する（パッケージによっては__version__属性が無い場合があります）
print(f"Spire.Doc Version: {getattr(spire.doc, '__version__', 'バージョン情報が見つかりません')}")

# インストール先を確認
print(f"Spire.Doc Installation Path: {os.path.dirname(spire.doc.__file__)}")
