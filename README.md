# MakeLanking

陸マガから記録を取得し、
大学別のランキングを作成してgoogle spreadsheetに出力するpythonスクリプト

## 概要

- seleniumで陸マガのランキングページにログイン
- beautifulsoupで記録をスクレイピング
- pandasで大学別ランキングを作成
- gspreadでgoogle spreadsheeetに出力

対象とする対校戦を指定可能
1. 京都インカレ
2. 関西インカレ
3. 七大戦
4. 同志社戦
5. 東大戦

---

## 必要環境
- Python 3.10+
- Google Sercvice Account

\\windows環境では今のところうまくいっていますが、WSL2ではまだ確認できてない

---

## インストール
```bash

# レポジトリをclone
git clone https://github.com/aoyg-y/MakeLanking.git

cd MakeLanking

# 仮想環境作成
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# ライブラリのインストール
pip install -r requirements.txt
```

---

## 環境変数設定

google APIのcredential jsonをディレクトリに配置\\
以下が参考
https://note.com/kohaku935/n/nf69f13012eb8

.envを作成\\
json keyと陸マガのアカウントを記入
```
EMAIL = xxxx
PASSWORD = xxxx
JSONKEY = xxxx
```

---

## 使い方

```python
ML = MakeLanking()
ML.make_lanking_spreadsheet(3,2024,"七大戦2024SB")
```

|引数|意味|
|-|-|
|competition|対校戦指定|
|year|記録を取得する年|
|ss_name|出力するスプレッドシート名|
