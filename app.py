import streamlit as st
from pdf2image import convert_from_bytes
import google.generativeai as genai
import json
import io
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment
from PIL import Image

# ── Gemini 設定 ───────────────────────────────────────────────────
genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
model = genai.GenerativeModel("gemini-2.0-flash")

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


# ── Geminiで1ページ分を解析 ───────────────────────────────────────
def analyze_receipt(img: Image.Image) -> dict:
    try:
        response = model.generate_content([PROMPT, img])
        text = response.text.strip()
        # コードブロックが含まれる場合に除去
        text = text.replace("```json", "").replace("```", "").strip()
        data = json.loads(text)
        return {
            "日付": data.get("date", ""),
            "店名": data.get("store", ""),
            "金額": data.get("amount"),
            "消費税8%": data.get("tax8"),
        }
    except Exception as e:
        return {"日付": "", "店名": f"読み取りエラー: {e}", "金額": None, "消費税8%": None}


# ── Excel生成 ─────────────────────────────────────────────────────
def build_excel(df: pd.DataFrame) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "現金領収書"

    headers = ["日付", "店名", "金額", "消費税8%"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    for _, row in df.iterrows():
        ws.append([
            row.get("日付", ""),
            row.get("店名", ""),
            int(row["金額"]) if pd.notna(row.get("金額")) and row.get("金額") != "" else None,
            int(row["消費税8%"]) if pd.notna(row.get("消費税8%")) and row.get("消費税8%") != "" else None,
        ])

    n = len(df)
    ws.cell(row=n + 2, column=1, value="合計").font = Font(bold=True)
    ws.cell(row=n + 2, column=3, value=f"=SUM(C2:C{n+1})")
    ws.cell(row=n + 2, column=4, value=f"=SUM(D2:D{n+1})")

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── メインUI ─────────────────────────────────────────────────────
st.set_page_config(page_title="現金領収書 → Excel", layout="wide")
st.title("現金領収書PDF → Excel 変換")
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

    filename = st.text_input("ファイル名", value="現金領収書一覧.xlsx")
    if not filename.endswith(".xlsx"):
        filename += ".xlsx"

    excel_bytes = build_excel(edited)
    st.download_button(
        label="Excelダウンロード",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
