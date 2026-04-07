[README.md](https://github.com/user-attachments/files/26524238/README.md)
# 現金領収書PDF → Excel 変換アプリ

スキャンした現金領収書のPDFをアップロードするだけで、日付・店名・金額・消費税8%をExcelに変換するWebアプリです。

**AI・APIキー不要。Python OCRのみで動作します。**

---

## 機能

- PDFをアップロード → 自動でOCR読み取り
- 日付・店名・金額・消費税8% を自動抽出
- 読み取り結果をブラウザ上で直接編集可能
- Excelファイル（.xlsx）としてダウンロード
- 和暦（令和・平成）→ 西暦に自動変換

---

## 使い方

1. アプリを開く
2. PDFファイルをアップロード
3. 読み取り結果を確認・修正
4. ファイル名を入力して「Excelダウンロード」

---

## ローカルで動かす場合

### 必要なもの

- Python 3.9以上
- Tesseract OCR（日本語対応）
- Poppler

### Tesseractのインストール

**Windows:**
[tesseract-ocr-w64-setup-5.x.exe](https://github.com/UB-Mannheim/tesseract/wiki) をインストール後、日本語データ（`jpn.traineddata`）を追加。

**Mac:**
```bash
brew install tesseract tesseract-lang
```

**Linux:**
```bash
sudo apt install tesseract-ocr tesseract-ocr-jpn poppler-utils
```

### セットアップ

```bash
git clone https://github.com/<your-username>/receipt-excel.git
cd receipt-excel
pip install -r requirements.txt
streamlit run app.py
```

---

## Streamlit Cloudへのデプロイ

1. このリポジトリをGitHubにプッシュ
2. [Streamlit Cloud](https://streamlit.io/cloud) でリポジトリを連携
3. Main file: `app.py` を指定してデプロイ

`packages.txt` により Tesseract と Poppler は自動インストールされます。

---

## 注意事項

- OCR精度はスキャン品質に依存します
- 手書き領収書は読み取り精度が下がる場合があります
- 読み取り後にテーブルで内容を確認・修正してからダウンロードしてください
