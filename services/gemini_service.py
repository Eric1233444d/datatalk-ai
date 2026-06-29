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
        "sample_rows": df.head(12).astype(str).to_dict(orient="records"),
        "column_summaries": {}
    }
    for col in df.columns:
        s = df[col]
        item = {
            "missing": int(s.isna().sum()),
            "dtype": str(s.dtype),
            "unique_count": int(s.nunique(dropna=False)),
            "examples": [str(x) for x in s.dropna().astype(str).unique()[:12]],
        }
        if s.nunique(dropna=False) <= 40:
            item["value_counts"] = s.fillna("空白").astype(str).value_counts().head(20).to_dict()
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


def _strip_value(value):
    value = str(value).strip().strip('「」『』"\'`“”‘’')
    value = re.sub(r"(的)?(資料|欄位|列|表|excel|xlsx|檔案|下載|另存|建一個|建立|產生|生成|匯出)$", "", value, flags=re.I).strip()
    return value


def parse_user_ops(user_text, df):
    """Gemini decides operations from real columns and samples. No fixed Excel schema is assumed."""
    client = get_client()
    if client is None or df is None:
        return None

    profile = make_data_profile(df)
    prompt = f"""
你是 DataTalk AI 的「通用 Excel Agent 指令規劃器」。
重點：不要假設固定欄位名稱。任何 Excel 都可能被上傳，你必須依據 columns、sample_rows、value_counts 判斷使用者想操作哪個欄位與資料。
請只輸出 JSON，不要輸出解釋、Markdown 或程式碼。

你可以輸出多個 operations，按順序執行。
允許的 type 與格式：
1. filter：篩選資料
{{"type":"filter","column":"必須是現有欄位","operator":"contains|not_contains|equals|not_equals|gt|gte|lt|lte|blank|not_blank","value":"值，可空字串"}}
2. search：全表搜尋
{{"type":"search","value":"值"}}
3. sort：排序
{{"type":"sort","column":"現有欄位","descending":false}}
4. count：統計分類數量
{{"type":"count","column":"現有欄位"}}
5. groupby：依欄位彙整
{{"type":"groupby","by":"現有欄位","target":"現有欄位或空字串","agg":"count|sum|mean|max|min"}}
6. keep_columns：只保留指定欄位
{{"type":"keep_columns","columns":["現有欄位1","現有欄位2"]}}
7. drop_columns：刪除指定欄位
{{"type":"drop_columns","columns":["現有欄位1"]}}
8. rename_column：修改欄位名稱
{{"type":"rename_column","old":"現有欄位","new":"新欄位名"}}
9. add_column：新增欄位，可用固定值或簡單規則文字
{{"type":"add_column","column":"新欄位名","value":"預設值或規則描述"}}
10. update_cells：更新符合條件的某欄值
{{"type":"update_cells","where_column":"現有欄位","operator":"contains|equals|not_contains|not_equals|blank|not_blank","where_value":"條件值","target_column":"現有欄位","new_value":"新值"}}
11. fill_blank：填補空白
{{"type":"fill_blank","column":"現有欄位或空字串代表全部","value":"填補值"}}
12. remove_duplicates：移除重複列
{{"type":"remove_duplicates","columns":["可空，空代表整列"]}}
13. summarize：產生摘要回答
{{"type":"summarize"}}
14. export：建立 Excel 下載
{{"type":"export"}}
15. ask_clarify：真的無法判斷時才用
{{"type":"ask_clarify","question":"你需要使用者補充的問題"}}
16. undo / redo
{{"type":"undo"}} 或 {{"type":"redo"}}

判斷規則：
- 使用者講「不是正常」、「非正常」、「排除正常」，operator 必須用 not_contains 或 not_equals，value 只放「正常」。
- 使用者講「異常」時，不要硬找固定欄位；請從資料摘要判斷哪些欄位和值像是狀態/備註/結果/問題，通常可用 filter not_contains 正常，或 search 異常/NG/失敗/待處理 等。
- 使用者只講某個值（例如公司、人名、編號），請從 sample/value_counts 找最可能的欄位；找不到就用 search。
- 使用者要求「修改、改成、更新」請用 update_cells 或 rename_column。
- 使用者要求「新增」請用 add_column。
- 使用者要求「彙整、統計、各...數量」請用 count 或 groupby。
- 使用者要求下載/匯出/另存/建立 Excel，最後加 export。
- 欄位名稱必須完全來自 columns，除非是 add_column.new column 或 rename_column.new。
- 如果可以合理執行，就不要 ask_clarify。

使用者需求：{user_text}

資料摘要 JSON：
{json.dumps(profile, ensure_ascii=False)}

只輸出：{{"operations":[...]}}
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
        allowed = {"filter","search","sort","count","groupby","keep_columns","drop_columns","rename_column","add_column","update_cells","fill_blank","remove_duplicates","summarize","export","ask_clarify","undo","redo"}
        cleaned = []
        for op in ops:
            if not isinstance(op, dict) or op.get("type") not in allowed:
                continue
            typ = op.get("type")
            op = dict(op)
            # Validate real columns
            if typ in {"filter","sort","count"} and str(op.get("column", "")) not in valid_cols:
                continue
            if typ == "groupby":
                if str(op.get("by", "")) not in valid_cols:
                    continue
                if op.get("target") and str(op.get("target")) not in valid_cols:
                    op["target"] = ""
            if typ in {"keep_columns","drop_columns","remove_duplicates"}:
                op["columns"] = [c for c in op.get("columns", []) if str(c) in valid_cols]
                if typ != "remove_duplicates" and not op["columns"]:
                    continue
            if typ == "rename_column" and str(op.get("old", "")) not in valid_cols:
                continue
            if typ == "update_cells":
                if str(op.get("where_column", "")) not in valid_cols or str(op.get("target_column", "")) not in valid_cols:
                    continue
            if typ == "fill_blank" and op.get("column") and str(op.get("column")) not in valid_cols:
                op["column"] = ""
            if typ in {"filter","search","update_cells"}:
                for key in ["value", "where_value", "new_value"]:
                    if key in op:
                        op[key] = _strip_value(op.get(key, ""))
            cleaned.append(op)
        return cleaned
    except Exception:
        return None


def ask_ai(user_text, df, tool_result=None):
    client = get_client()
    if client is None:
        return "目前尚未設定 GEMINI_API_KEY。請確認 Streamlit Cloud Secrets。"
    profile = make_data_profile(df)
    prompt = f"""
你是 DataTalk AI，一個 Excel 資料分析與整理助理。請用繁體中文回答。
不能亂編資料；請根據資料摘要與工具結果回答。

使用者問題：{user_text}
資料摘要：{json.dumps(profile, ensure_ascii=False)}
工具結果：{tool_result if tool_result else "尚無"}
"""
    try:
        res = client.models.generate_content(model=MODEL_NAME, contents=prompt)
        return res.text
    except Exception as e:
        return f"Gemini 回覆失敗，但 Excel 基本操作仍可使用。錯誤：{e}"
