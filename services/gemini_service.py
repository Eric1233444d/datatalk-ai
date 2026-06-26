import json
import os
import re
import streamlit as st
from google import genai

MODEL_NAME = "gemini-2.5-flash"


def get_api_key():
    try:
        key = st.secrets.get("GEMINI_API_KEY", "")
    except Exception:
        key = ""
    return key or os.getenv("GEMINI_API_KEY", "")


def get_client():
    api_key = get_api_key()
    if not api_key:
        return None
    return genai.Client(api_key=api_key)


def make_data_profile(df):
    profile = {
        "rows": len(df),
        "columns": list(map(str, df.columns)),
        "missing_total": int(df.isna().sum().sum()),
        "duplicate_rows": int(df.duplicated().sum()),
        "sample_rows": df.head(8).astype(str).to_dict(orient="records"),
        "column_summaries": {}
    }
    for col in df.columns:
        s = df[col]
        item = {
            "missing": int(s.isna().sum()),
            "dtype": str(s.dtype),
            "unique_count": int(s.nunique(dropna=False)),
        }
        if s.nunique(dropna=False) <= 30:
            item["value_counts"] = s.fillna("空白").astype(str).value_counts().head(15).to_dict()
        if str(s.dtype).startswith(("int", "float")):
            item["sum"] = float(s.sum(skipna=True))
            item["mean"] = float(s.mean(skipna=True))
            item["min"] = float(s.min(skipna=True))
            item["max"] = float(s.max(skipna=True))
        profile["column_summaries"][str(col)] = item
    return profile


def _extract_json(text):
    if not text:
        return None
    text = text.strip()
    fenced = re.search(r"```(?:json)?\s*(.*?)\s*```", text, re.S | re.I)
    if fenced:
        text = fenced.group(1).strip()
    start = text.find("{")
    end = text.rfind("}")
    if start >= 0 and end > start:
        text = text[start:end + 1]
    try:
        return json.loads(text)
    except Exception:
        return None


def parse_user_ops(user_text, df):
    """Ask Gemini to convert natural language into safe Excel operations.
    If Gemini is unavailable or returns invalid JSON, return None and let local parser handle it.
    """
    client = get_client()
    if client is None or df is None:
        return None

    profile = make_data_profile(df)
    prompt = f"""
你是 DataTalk AI 的 Excel 指令解析器。請只輸出 JSON，不要輸出解釋。
你的任務：把使用者中文需求轉成可執行 operations。

允許的 type：
- filter: 篩選某欄包含或不包含某值。欄位必須來自 columns。格式：{{"type":"filter","column":"欄位名","value":"值","negative":false}}
- search: 全表搜尋。格式：{{"type":"search","value":"值"}}
- abnormal: 搜尋異常資料。格式：{{"type":"abnormal"}}
- count: 統計欄位數量。格式：{{"type":"count","column":"欄位名"}}
- sort: 排序。格式：{{"type":"sort","column":"欄位名","descending":false}}
- add_column: 新增欄位。格式：{{"type":"add_column","column":"欄位名","value":"預設值"}}
- fill_blank: 填補空白。格式：{{"type":"fill_blank","value":"值"}}
- export: 建立/下載 Excel。格式：{{"type":"export"}}
- undo: 復原。格式：{{"type":"undo"}}
- redo: 重做。格式：{{"type":"redo"}}

重要規則：
1. 使用者說「不是正常」、「不包含正常」、「排除正常」時，value 必須是「正常」，negative 必須是 true。
2. 不要把「的欄位」、「的資料」、「下載」、「建新的表」放進 value。
3. 如果使用者說「華新工程」且資料摘要中看起來是供應商，請用供應商欄位 filter。
4. 一句話有多步驟就輸出多個 operations，例如先 filter 再 sort 再 export。
5. 如果使用者要求 Excel/下載/匯出/另存/建立新表，最後加 export。
6. 欄位名稱必須完全使用 columns 中存在的文字。
7. 沒有明確可執行操作時，輸出空陣列。

使用者需求：{user_text}

資料摘要 JSON：
{json.dumps(profile, ensure_ascii=False)}

只輸出：
{{"operations":[...]}}
"""
    try:
        res = client.models.generate_content(model=MODEL_NAME, contents=prompt)
        data = _extract_json(getattr(res, "text", ""))
        if not isinstance(data, dict):
            return None
        ops = data.get("operations")
        if not isinstance(ops, list):
            return None
        valid_cols = set(map(str, df.columns))
        cleaned = []
        allowed = {"filter", "search", "abnormal", "count", "sort", "add_column", "fill_blank", "export", "undo", "redo"}
        for op in ops:
            if not isinstance(op, dict) or op.get("type") not in allowed:
                continue
            op = dict(op)
            if op.get("type") in {"filter", "count", "sort"}:
                if str(op.get("column", "")) not in valid_cols:
                    continue
            if op.get("type") == "filter":
                value = str(op.get("value", "")).strip().strip('「」『』"\'`“”‘’')
                value = re.sub(r"(的)?(資料|欄位|列|表|excel|檔案|下載|另存|建一個|建立|產生|生成)$", "", value, flags=re.I).strip()
                if not value:
                    continue
                op["value"] = value
                op["negative"] = bool(op.get("negative", False))
            if op.get("type") == "search" and not str(op.get("value", "")).strip():
                continue
            cleaned.append(op)
        return cleaned
    except Exception:
        return None


def ask_ai(user_text, df, tool_result=None):
    client = get_client()
    if client is None:
        return "目前尚未設定 GEMINI_API_KEY。請確認 .streamlit/secrets.toml 或 Streamlit Cloud Secrets。"

    profile = make_data_profile(df)
    prompt = f"""
你是 DataTalk AI，一個 Excel 資料分析與整理助理。
請用繁體中文、口語但專業的方式回答。
你不能亂編資料。若資料不足，請說明需要哪些欄位。
若工具已經執行，請根據工具結果說明下一步建議。

使用者問題：
{user_text}

資料摘要 JSON：
{json.dumps(profile, ensure_ascii=False)}

工具執行結果：
{tool_result if tool_result else "尚無"}
"""
    try:
        res = client.models.generate_content(model=MODEL_NAME, contents=prompt)
        return res.text
    except Exception as e:
        return f"Gemini 回覆失敗，但 Excel 基本操作仍可使用。錯誤：{e}"
