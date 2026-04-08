import streamlit as st
from pdf2image import convert_from_bytes
from groq import Groq
import base64
import json
import io
import time
import pandas as pd
from PIL import Image

# ── Groq 設定 ─────────────────────────────────────────────────────
client = Groq(api_key=st.secrets["GROQ_API_KEY"])

PROMPT = """この画像は日本の領収書またはレシートです。以下の情報を抽出してJSONで返してください。

{
  "date": "YYYY/MM/DD形式の日付（和暦は西暦に変換。令和7年=2025年、令和6年=2024年、平成31年=2019年）",
  "store": "店名または発行者名",
  "amount": 合計金額（税込・数値のみ・カンマなし）,
  "tax8": 消費税8%軽減税率分の税額（数値のみ。記載がない場合はnull）
}

注意：
- 金額が読み取れない場合は null
- 日付・店名が読み取れない場合は ""
- JSONのみ返すこと（説明文・コードブロック不要）"""


# ── 画像をbase64に変換 ───────────────────────────────────────────
def img_to_base64(img: Image.Image) -> str:
    buf = io.BytesIO()
    img.save(buf, format="JPEG")
    return base64.b64encode(buf.getvalue()).decode()


# ── Groqで1ページ分を解析（リトライあり） ────────────────────────
def analyze_receipt(img: Image.Image, retries: int = 3) -> dict:
    b64 = img_to_base64(img)
    for attempt in range(retries):
        try:
            response = client.chat.completions.create(
                model="meta-llama/llama-4-scout-17b-16e-instruct",
                messages=[{
                    "role": "user",
                    "content": [
                        {"type": "text", "text": PROMPT},
                        {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}}
                    ]
                }]
            )
            text = response.choices[0].message.content.strip()
            text = text.replace("```json", "").replace("```", "").strip()
            data = json.loads(text)
            return {
                "日付": data.get("date", ""),
                "店名": data.get("store", ""),
                "金額": data.get("amount"),
                "消費税8%": data.get("tax8"),
            }
        except Exception as e:
            if "429" in str(e) and attempt < retries - 1:
                wait = 60 * (attempt + 1)
                st.warning(f"リクエスト制限中... {wait}秒待ってリトライします ({attempt+1}/{retries})")
                time.sleep(wait)
            else:
                return {"日付": "", "店名": f"読み取りエラー: {e}", "金額": None, "消費税8%": None}
    return {"日付": "", "店名": "読み取り失敗", "金額": None, "消費税8%": None}


# ── CSV生成 ─────────────────────────────────────────────────────
def build_csv(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    df.to_csv(buf, index=False, encoding='utf-8-sig')
    return buf.getvalue()


# ── メインUI ─────────────────────────────────────────────────────
st.set_page_config(page_title="現金領収書 → CSV", layout="wide")
st.title("現金領収書PDF → CSV 変換")
st.caption("スキャンしたPDFをアップロードすると、日付・店名・金額・消費税8%を読み取ります。")

uploaded = st.file_uploader("PDFファイルをアップロード", type="pdf")

if uploaded:
    with st.spinner("読み取り中..."):
        try:
            images = convert_from_bytes(uploaded.read(), dpi=200)
        except Exception as e:
            st.error(f"PDF変換エラー: {e}")
            st.stop()

        records = []
        bar = st.progress(0, text="解析中...")
        for i, img in enumerate(images):
            record = analyze_receipt(img)
            records.append(record)
            bar.progress((i + 1) / len(images), text=f"{i+1} / {len(images)} ページ完了")
        bar.empty()

    st.success(f"{len(records)} ページを読み取りました")

    df = pd.DataFrame(records)

    st.subheader("読み取り結果（直接編集できます）")
    edited = st.data_editor(
        df,
        num_rows="dynamic",
        column_config={
            "日付": st.column_config.TextColumn("日付", width=120),
            "店名": st.column_config.TextColumn("店名", width=300),
            "金額": st.column_config.NumberColumn("金額", format="%d"),
            "消費税8%": st.column_config.NumberColumn("消費税8%", format="%d"),
        },
        use_container_width=True,
    )

    filename = st.text_input("ファイル名", value="現金領収書一覧.csv")
    if not filename.endswith(".csv"):
        filename += ".csv"

    csv_bytes = build_csv(edited)
    st.download_button(
        label="CSVダウンロード",
        data=csv_bytes,
        file_name=filename,
        mime="text/csv",
    )
