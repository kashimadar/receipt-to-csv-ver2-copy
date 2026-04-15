# 現金領収書PDF → CSV 変換アプリ

スキャンした現金領収書のPDFをアップロードするだけで、日付・店名・金額・消費税8%をCSVに変換するWebアプリです。

**Groq API（llama-4-scout）または Google Gemini API を選択して使用できます。**

---

## 機能

- PDFをアップロード → 自動でAI読み取り
- **GeminiとGroqをUIで切り替え可能**
- 日付・店名・金額・消費税8% を自動抽出
- 読み取り結果をブラウザ上で直接編集可能
- CSVファイル（.csv）としてダウンロード
- 和暦（令和・平成）→ 西暦に自動変換

---

## 使い方

1. アプリを開く
2. PDFファイルをアップロード
3. 読み取り結果を確認・修正
4. ファイル名を入力して「CSVダウンロード」

---

## ローカルで動かす場合

### 必要なもの

- Python 3.9以上
- Poppler
- Groq APIキー（Groq使用時）
- Google Gemini APIキー（Gemini使用時）

### Popplerのインストール

**Windows:**
[Poppler for Windows](https://github.com/oschwartz10612/poppler-windows/releases) をインストールし、`bin` フォルダにパスを通す。

**Mac:**
```bash
brew install poppler
```

**Linux:**
```bash
sudo apt install poppler-utils
```

### セットアップ

```bash
git clone https://github.com/<your-username>/receipt-to-csv-ver2-copy.git
cd receipt-to-csv-ver2-copy
pip install -r requirements.txt
```

`.streamlit/secrets.toml` を作成して APIキー を設定:

```toml
GROQ_API_KEY = "your-groq-api-key"
bankpass = "your-gemini-api-key"
```

```bash
streamlit run app.py
```

---

## Streamlit Cloudへのデプロイ

1. このリポジトリをGitHubにプッシュ
2. [Streamlit Cloud](https://streamlit.io/cloud) でリポジトリを連携
3. Main file: `app.py` を指定してデプロイ
4. Secrets に `GROQ_API_KEY` と `bankpass`（Gemini APIキー）を設定

`packages.txt` により Poppler は自動インストールされます。

---

## 注意事項

- AI読み取り精度はスキャン品質に依存します
- 手書き領収書は読み取り精度が下がる場合があります
- 読み取り後にテーブルで内容を確認・修正してからダウンロードしてください
- Groq APIの無料枠にはリクエスト制限があります（制限時は自動リトライします）
- Gemini使用時はGemini APIの利用制限に準じます
