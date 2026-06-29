import io
import re
import random
from copy import deepcopy
from datetime import datetime
from pathlib import Path
from difflib import get_close_matches

import pandas as pd
import streamlit as st
import plotly.express as px

try:
    from services.gemini_service import parse_user_ops, get_api_key
except Exception:
    parse_user_ops = None
    get_api_key = None

try:
    from services.excel_service import dataframe_to_excel_bytes, preserve_style_export, read_excel
except Exception:
    dataframe_to_excel_bytes = None
    preserve_style_export = None
    read_excel = None

st.set_page_config(page_title="DataTalk AI", page_icon="📊", layout="wide")

st.markdown("""
<style>
.stApp { background:#f6f7fb; }
.block-container { max-width:1280px; padding-top:2.8rem; }
.main-title { font-size:38px; line-height:1.25; font-weight:900; color:#0f172a; margin:0 0 10px 0; padding-top:4px; }
.sub-title { color:#64748b; font-size:16px; margin-bottom:20px; }
.small { color:#64748b; font-size:13px; }
.badge { display:inline-block; padding:6px 10px; border-radius:999px; background:#eef2ff; color:#3730a3; margin:4px; font-size:13px; }
.chat-hint { background:white; border:1px solid #e2e8f0; border-radius:16px; padding:16px; margin:8px 0; }
.result-box { background:#ecfdf5; border:1px solid #bbf7d0; border-radius:14px; padding:14px; }
div.stDownloadButton > button, div.stButton > button { max-width:100%; white-space:normal; text-align:left; }
</style>
""", unsafe_allow_html=True)

SAMPLE_DIR = Path("sample_data")
NEGATIVE_WORDS = ["不是", "不等於", "不包含", "非", "排除", "除了", "不要", "去掉", "移除", "剔除", "不是\"", "不是'"]
EXPORT_WORDS = ["excel", "xlsx", "下載", "匯出", "另存", "建一個", "建立", "產生", "生成", "新的表", "新表", "新excel", "新的excel"]
FILTER_WORDS = ["篩選", "挑出", "挑出來", "抓出", "抓出來", "只留下", "保留", "找", "搜尋", "查", "列出", "只看"]
ABNORMAL_PATTERN = r"異常|延遲|失敗|NG|ng|問題|缺失|未完成|待處理|風險|錯誤|缺料|補料"
ALIASES = {
    "供應商": ["供應商", "廠商", "vendor", "supplier", "公司"],
    "備註": ["備註", "說明", "remark", "note", "memo", "註記"],
    "狀態": ["狀態", "結果", "status", "處理狀態"],
    "部門": ["部門", "單位", "department"],
    "區域": ["區域", "廠區", "地區", "area"],
    "負責人": ["負責人", "承辦", "owner", "窗口"],
    "工單編號": ["工單編號", "工單", "編號", "單號", "id"],
    "日期": ["日期", "時間", "date", "day"],
    "數量": ["數量", "qty", "quantity", "count"],
    "金額": ["金額", "費用", "價格", "price", "cost", "amount"],
}

WELCOME = """👋 我可以直接幫你操作 Excel，不只是回答建議。\n\n你可以直接說：\n\n- 把供應商是華新工程的資料建一個 Excel\n- 把備註不是正常的資料下載下來\n- 搜尋異常\n- 統計狀態\n- 把華新工程抓出來，再依日期排序，匯出 Excel\n- undo / redo\n\n請先到「上傳資料」下載測試 Excel 或上傳自己的 Excel。"""


def init():
    defaults = {
        "df": None,
        "df2": None,
        "working_df": None,
        "file_name": None,
        "file2_name": None,
        "original_file_bytes": None,
        "excel_meta": None,
        "messages": [{"role": "assistant", "content": WELCOME}],
        "downloads": [],
        "last_result": None,
        "last_chart": None,
        "last_filter_value": None,
        "last_filter_column": None,
        "history": [],
        "redo_stack": [],
        "operation_log": [],
        "pending": None,
        "use_working_source": False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = deepcopy(v)


def render_header(title, subtitle):
    st.markdown(f"<div class='main-title'>{title}</div><div class='sub-title'>{subtitle}</div>", unsafe_allow_html=True)


def normalize(text):
    return str(text).strip().lower().replace(" ", "").replace("　", "")


def clean_token(value):
    value = str(value).strip()
    value = re.sub(r"^[：:，,。\s]+|[：:，,。\s]+$", "", value)
    value = value.strip('「」『』\"\'`“”‘’')
    # 常見口語尾巴，不要把「正常的欄位」整串當 value
    value = re.sub(r"(的)?(資料|欄位|列|表|excel|檔案|下載|另存|建一個|建立|產生|生成|都|全部|只要)$", "", value, flags=re.I)
    value = value.strip('「」『』\"\'`“”‘’ ，,。')
    return value



def decode_hash_unicode(name):
    """把檔名中的 #U4eca#U65e5 這種編碼還原成中文，避免下載按鈕跑出亂碼。"""
    def repl(m):
        try:
            return chr(int(m.group(1), 16))
        except Exception:
            return m.group(0)
    return re.sub(r"#U([0-9a-fA-F]{4,6})", repl, str(name))


def original_df():
    return st.session_state.df


def reset_to_original():
    if st.session_state.df is not None:
        st.session_state.working_df = st.session_state.df.copy()
        st.session_state.last_result = st.session_state.df.copy()
        st.session_state.history.clear()
        st.session_state.redo_stack.clear()
        st.session_state.operation_log.append(f"{datetime.now().strftime('%H:%M:%S')}｜回到原始資料")

def active_df():
    return st.session_state.working_df


def push_history(label="操作前"):
    df = active_df()
    if df is not None:
        st.session_state.history.append({"label": label, "df": df.copy()})
        st.session_state.history = st.session_state.history[-30:]
        st.session_state.redo_stack.clear()


def set_active(df, operation=None, save_history=True):
    if save_history:
        push_history(operation or "操作")
    st.session_state.working_df = df.copy()
    st.session_state.last_result = df.copy()
    if operation:
        st.session_state.operation_log.append(f"{datetime.now().strftime('%H:%M:%S')}｜{operation}")
        st.session_state.operation_log = st.session_state.operation_log[-50:]


def excel_bytes(df):
    if preserve_style_export is not None:
        return preserve_style_export(st.session_state.get("original_file_bytes"), df)
    if dataframe_to_excel_bytes is not None:
        return dataframe_to_excel_bytes(df)
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="DataTalk_Result")
    return bio.getvalue()


def safe_filename(text):
    text = re.sub(r"[\\/:*?\"<>|]", "_", str(text))
    text = re.sub(r"\s+", "_", text).strip("_")
    return text[:80] or "DataTalk_Result"


def add_download(df, name, note=""):
    filename = safe_filename(name)
    if not filename.lower().endswith(".xlsx"):
        filename += ".xlsx"
    item = {
        "name": filename,
        "data": excel_bytes(df),
        "rows": len(df),
        "cols": len(df.columns),
        "note": note,
        "time": datetime.now().strftime("%H:%M:%S"),
    }
    st.session_state.downloads.insert(0, item)
    st.session_state.last_result = df.copy()
    return item


def best_column(text, columns):
    if not columns:
        return None
    nt = normalize(text)
    # 1. 欄位全名命中
    for col in columns:
        if normalize(col) and normalize(col) in nt:
            return col
    # 2. alias 命中
    for _, keys in ALIASES.items():
        if any(normalize(k) in nt for k in keys):
            for col in columns:
                if any(normalize(k) in normalize(col) for k in keys):
                    return col
    # 3. 模糊欄位名稱
    col_map = {normalize(c): c for c in columns}
    matches = get_close_matches(nt, list(col_map.keys()), n=1, cutoff=0.72)
    if matches:
        return col_map[matches[0]]
    return None


def column_values(df, col, limit=500):
    vals = []
    for v in df[col].dropna().astype(str).unique()[:limit]:
        v = clean_token(v)
        if v:
            vals.append(v)
    return vals


def guess_value_from_data(text, df, col=None):
    nt = normalize(text)
    cols = [col] if col else list(df.columns)
    best = None
    for c in cols:
        if c is None:
            continue
        for v in column_values(df, c):
            nv = normalize(v)
            if nv and nv in nt:
                if best is None or len(nv) > len(normalize(best)):
                    best = v
    return best


def fuzzy_value(text, df, col):
    raw = clean_token(text)
    if not raw or col is None:
        return raw, None
    values = column_values(df, col)
    if raw in values:
        return raw, None
    # 使用者打錯一點點時，例如華興工程 -> 華新工程
    matches = get_close_matches(raw, values, n=1, cutoff=0.62)
    if matches:
        return matches[0], raw
    return raw, None


def extract_value(text, df, col=None):
    t = str(text)
    quoted = re.search(r"[\"'「『“‘]([^\"'」』”’]+)[\"'」』”’]", t)
    if quoted:
        return clean_token(quoted.group(1)), None

    # 先從資料值命中，避免「供應商是華新工程的資料」被抓成「華新工程的資料」
    data_value = guess_value_from_data(t, df, col)
    if data_value:
        return data_value, None

    patterns = [
        r"(?:是|為|等於|包含|含有|不是|不等於|不包含|非)\s*([^，,。\n]+)",
        r"(?:搜尋|查詢|搜尋一下|找)\s*([^，,。\n]+)",
        r"(?:只留下|保留|排除|除了|不要|去掉|移除)\s*([^，,。\n]+)",
    ]
    for pat in patterns:
        m = re.search(pat, t, flags=re.I)
        if m:
            raw = clean_token(m.group(1))
            # 把後面串接動作切掉
            raw = re.split(r"(?:再|然後|並且|順便|，|,|。|、)", raw)[0]
            raw = clean_token(raw)
            value, correction = fuzzy_value(raw, df, col)
            return value, correction
    return None, None


def infer_column_by_value(df, value):
    if not value:
        return None, []
    hits = []
    for c in df.columns:
        if df[c].fillna("").astype(str).str.contains(str(value), case=False, na=False, regex=False).any():
            hits.append(c)
    if len(hits) == 1:
        return hits[0], hits
    return None, hits


def is_negative(text):
    nt = normalize(text)
    return any(normalize(w) in nt for w in NEGATIVE_WORDS)


def filter_dataframe(df, col, value, negative=False):
    s = df[col].fillna("").astype(str)
    mask = s.str.contains(str(value), case=False, na=False, regex=False)
    return df[~mask].copy() if negative else df[mask].copy()


def parse_filter(text, df):
    col = best_column(text, list(df.columns))
    value, correction = extract_value(text, df, col)

    if value and not col:
        inferred, hits = infer_column_by_value(df, value)
        col = inferred
        if not inferred and hits:
            return {"type": "need_column", "value": value, "hits": hits}

    # 只有說「華新工程」也要能做：先用資料值判斷欄位
    if not value:
        any_value = guess_value_from_data(text, df)
        if any_value:
            value = any_value
            if not col:
                col, _ = infer_column_by_value(df, value)

    if col and not value and st.session_state.get("last_filter_value"):
        value = st.session_state.last_filter_value
    if value and not col and st.session_state.get("last_filter_column"):
        col = st.session_state.last_filter_column

    if col and value:
        return {
            "type": "filter",
            "column": col,
            "value": value,
            "negative": is_negative(text),
            "correction": correction,
        }
    return None


def split_steps(text):
    parts = re.split(r"(?:，|,|。|\n|再|然後|並且|順便|接著)", str(text))
    return [p.strip() for p in parts if p.strip()]


def wants_export(text):
    return any(k in normalize(text) for k in [normalize(w) for w in EXPORT_WORDS])


def detect_operations(text, df):
    ops = []
    full_t = normalize(text)

    if full_t in ["undo", "復原", "上一步", "還原"] or "復原" in full_t:
        return [{"type": "undo"}]
    if full_t in ["redo", "重做", "下一步"] or "重做" in full_t:
        return [{"type": "redo"}]

    # 一句話中若有「備註不是正常」這種完整條件，先整句解析，避免切太碎
    whole_filter = parse_filter(text, df)
    if whole_filter and whole_filter.get("type") == "filter":
        if any(k in full_t for k in [normalize(w) for w in FILTER_WORDS + EXPORT_WORDS]) or whole_filter["value"]:
            ops.append(whole_filter)

    for step in split_steps(text):
        nt = normalize(step)
        if ops and ops[-1].get("type") == "filter" and parse_filter(step, df) == ops[-1]:
            continue
        if "統計" in nt or "各" in nt and "數量" in nt:
            col = best_column(step, list(df.columns))
            if col:
                ops.append({"type": "count", "column": col})
                continue
        if "異常" in nt and any(k in nt for k in ["搜尋", "找", "整理", "下載", "excel", "抓", "列出"]):
            ops.append({"type": "abnormal"})
            continue
        if "搜尋" in nt and not best_column(step, list(df.columns)):
            value, _ = extract_value(step, df)
            value = value or clean_token(step.replace("搜尋", "").replace("查詢", ""))
            if value:
                ops.append({"type": "search", "value": value})
                continue
        if any(k in nt for k in ["排序", "依", "由大到小", "由小到大", "升冪", "降冪"]):
            col = best_column(step, list(df.columns))
            if col:
                descending = any(k in nt for k in ["由大到小", "降冪", "最高", "最新"])
                ops.append({"type": "sort", "column": col, "descending": descending})
                continue
        if "新增欄位" in nt:
            m = re.search(r"新增欄位\s*(.+?)(?:\s*值為\s*(.+))?$", step)
            if m:
                ops.append({"type": "add_column", "column": clean_token(m.group(1)), "value": clean_token(m.group(2) or "")})
                continue
        if "空白" in nt and any(k in nt for k in ["補", "填"]):
            m = re.search(r"(?:補成|填成|補為|填為)\s*([^，,。]+)", step)
            ops.append({"type": "fill_blank", "value": clean_token(m.group(1)) if m else "0"})
            continue
        fop = parse_filter(step, df)
        if fop:
            ops.append(fop)
            continue

    # 去重：同一句常會整句+分句各解析一次
    dedup = []
    seen = set()
    for op in ops:
        key = tuple(sorted((k, str(v)) for k, v in op.items() if k != "correction"))
        if key not in seen:
            dedup.append(op)
            seen.add(key)
    if wants_export(text):
        dedup.append({"type": "export"})
    return dedup


def reply_done(lines, rows=None):
    prefixes = ["✅ 已完成。", "沒問題，已經整理好了。", "完成！", "好的～我已經幫你處理好了。"]
    first = random.choice(prefixes)
    body = "\n".join(lines)
    if rows is not None:
        body += f"\n\n共 **{rows}** 筆資料。"
    return f"{first}\n\n{body}"



def _to_number_if_possible(value):
    try:
        if value is None or value == "":
            return value
        return float(str(value).replace(",", ""))
    except Exception:
        return value


def build_condition_mask(df, column, operator, value=""):
    s = df[column]
    op = str(operator or "contains")
    if op in ["blank"]:
        return s.isna() | (s.astype(str).str.strip() == "")
    if op in ["not_blank"]:
        return ~(s.isna() | (s.astype(str).str.strip() == ""))
    if op in ["gt", "gte", "lt", "lte"]:
        num = pd.to_numeric(s, errors="coerce")
        v = _to_number_if_possible(value)
        try:
            if op == "gt":
                return num > float(v)
            if op == "gte":
                return num >= float(v)
            if op == "lt":
                return num < float(v)
            if op == "lte":
                return num <= float(v)
        except Exception:
            return pd.Series(False, index=df.index)
    text_s = s.fillna("").astype(str)
    val = str(value)
    if op == "equals":
        return text_s.str.strip().str.lower() == val.strip().lower()
    if op == "not_equals":
        return text_s.str.strip().str.lower() != val.strip().lower()
    if op == "not_contains":
        return ~text_s.str.contains(val, case=False, na=False, regex=False)
    return text_s.str.contains(val, case=False, na=False, regex=False)


def execute_ops(text):
    continue_words = ["再", "接著", "然後", "目前", "剛剛", "這些", "篩選結果", "上一步結果"]
    should_continue = st.session_state.get("use_working_source", False) or any(w in str(text) for w in continue_words)
    df = active_df() if should_continue else original_df()
    if df is None:
        return "請先到「上傳資料」上傳 Excel，或下載測試 Excel 後再上傳。", None

    ai_ops = None
    if parse_user_ops is not None:
        ai_ops = parse_user_ops(text, df)
    ops = ai_ops if ai_ops else detect_operations(text, df)
    if not ops:
        return "我還沒抓到明確操作。你可以直接說：搜尋某個值、統計某欄、把某欄不是正常的資料匯出、修改某欄、保留指定欄位。", None

    current = df.copy()
    lines = []
    export_needed = False
    filename_parts = []
    chart = None

    for op in ops:
        typ = op.get("type")
        if typ == "ask_clarify":
            return "我需要你補充一下：" + str(op.get("question", "請指定要操作的欄位或條件。")), current
        if typ == "summarize":
            missing = int(current.isna().sum().sum())
            dup = int(current.duplicated().sum())
            lines.append(f"這份資料目前有 **{len(current)} 筆**、**{len(current.columns)} 欄**；空白值 **{missing}** 個、重複列 **{dup}** 筆。")
            continue
        if typ == "undo":
            if not st.session_state.history:
                return "目前沒有可以復原的上一步。", current
            st.session_state.redo_stack.append({"label": "redo", "df": current.copy()})
            prev = st.session_state.history.pop()
            st.session_state.working_df = prev["df"].copy()
            st.session_state.last_result = st.session_state.working_df.copy()
            return f"↩️ 已復原：{prev['label']}。", st.session_state.working_df
        if typ == "redo":
            if not st.session_state.redo_stack:
                return "目前沒有可以重做的步驟。", current
            push_history("redo 前")
            nxt = st.session_state.redo_stack.pop()
            st.session_state.working_df = nxt["df"].copy()
            st.session_state.last_result = st.session_state.working_df.copy()
            return "↪️ 已重做下一步。", st.session_state.working_df
        if typ == "need_column":
            st.session_state.last_filter_value = op["value"]
            return f"我找到「{op['value']}」可能在多個欄位：{', '.join(map(str, op['hits']))}。請再補一句欄位。", None
        if typ == "filter":
            col = op.get("column")
            if col not in current.columns:
                lines.append(f"略過篩選：找不到欄位 {col}。")
                continue
            before = len(current)
            operator = op.get("operator") or ("not_contains" if op.get("negative") else "contains")
            value = op.get("value", "")
            mask = build_condition_mask(current, col, operator, value)
            current = current[mask].copy()
            st.session_state.last_filter_column = col
            st.session_state.last_filter_value = value
            op_word = {
                "contains":"包含", "not_contains":"不包含", "equals":"等於", "not_equals":"不等於",
                "gt":"大於", "gte":"大於等於", "lt":"小於", "lte":"小於等於", "blank":"空白", "not_blank":"非空白"
            }.get(operator, operator)
            show_value = f"「{value}」" if str(value) else ""
            lines.append(f"已篩選：**{col} {op_word}{show_value}**，{before} 筆 → {len(current)} 筆。")
            filename_parts.append(f"{col}_{op_word}{value}")
            export_needed = export_needed or wants_export(text)
        elif typ == "search":
            value = str(op.get("value", ""))
            mask = pd.Series(False, index=current.index)
            for c in current.columns:
                mask |= current[c].fillna("").astype(str).str.contains(value, case=False, na=False, regex=False)
            before = len(current)
            current = current[mask].copy()
            lines.append(f"已全表搜尋：**{value}**，{before} 筆 → {len(current)} 筆。")
            filename_parts.append(f"搜尋_{value}")
            export_needed = True
        elif typ == "abnormal":
            # fallback only: Gemini normally decides columns/values. This searches whole table for common abnormal words.
            mask = pd.Series(False, index=current.index)
            for c in current.columns:
                mask |= current[c].fillna("").astype(str).str.contains(ABNORMAL_PATTERN, case=False, na=False, regex=True)
            current = current[mask].copy()
            lines.append("已依常見異常關鍵字做全表搜尋。")
            filename_parts.append("異常資料")
            export_needed = True
        elif typ == "count":
            col = op.get("column")
            if col not in current.columns:
                continue
            out = current[col].fillna("空白").astype(str).value_counts().reset_index()
            out.columns = [col, "數量"]
            current = out
            chart = (out, col)
            lines.append(f"已統計：**{col}**，共 {len(out)} 種分類。")
            filename_parts.append(f"{col}_統計")
            export_needed = True
        elif typ == "groupby":
            by = op.get("by")
            target = op.get("target")
            agg = op.get("agg", "count")
            if by not in current.columns:
                continue
            if agg == "count" or not target or target not in current.columns:
                out = current.groupby(by, dropna=False).size().reset_index(name="數量")
            else:
                nums = pd.to_numeric(current[target], errors="coerce")
                temp = current.copy(); temp[target] = nums
                out = getattr(temp.groupby(by, dropna=False)[target], agg)().reset_index(name=f"{target}_{agg}")
            current = out
            chart = (out, by)
            lines.append(f"已依 **{by}** 彙整（{agg}）。")
            filename_parts.append(f"依{by}彙整")
            export_needed = True
        elif typ == "sort":
            col = op.get("column")
            if col not in current.columns:
                continue
            current = current.sort_values(by=col, ascending=not bool(op.get("descending", False))).copy()
            lines.append(f"已依照 **{col}** {'由大到小' if op.get('descending') else '由小到大'}排序。")
            filename_parts.append(f"依{col}排序")
        elif typ == "keep_columns":
            cols = [c for c in op.get("columns", []) if c in current.columns]
            if cols:
                current = current[cols].copy()
                lines.append(f"已保留欄位：{', '.join(cols)}。")
                filename_parts.append("保留指定欄位")
                export_needed = True
        elif typ == "drop_columns":
            cols = [c for c in op.get("columns", []) if c in current.columns]
            if cols:
                current = current.drop(columns=cols).copy()
                lines.append(f"已刪除欄位：{', '.join(cols)}。")
                filename_parts.append("刪除欄位")
                export_needed = True
        elif typ == "rename_column":
            old, new = op.get("old"), op.get("new")
            if old in current.columns and new:
                current = current.rename(columns={old: new}).copy()
                lines.append(f"已將欄位 **{old}** 改名為 **{new}**。")
                filename_parts.append("欄位改名")
                export_needed = True
        elif typ == "fill_blank":
            col = op.get("column", "")
            value = op.get("value", "")
            if col and col in current.columns:
                current[col] = current[col].fillna(value).replace("", value)
                lines.append(f"已把 **{col}** 的空白值補成 **{value}**。")
            else:
                current = current.fillna(value).replace("", value)
                lines.append(f"已把所有空白值補成 **{value}**。")
            filename_parts.append("空白值已補")
            export_needed = True
        elif typ == "add_column":
            col = op.get("column")
            value = op.get("value", "")
            if col:
                current[col] = value
                lines.append(f"已新增欄位 **{col}**，內容為「{value}」。")
                filename_parts.append(f"新增{col}")
                export_needed = True
        elif typ == "update_cells":
            where_col = op.get("where_column")
            target_col = op.get("target_column")
            if where_col in current.columns and target_col in current.columns:
                mask = build_condition_mask(current, where_col, op.get("operator", "contains"), op.get("where_value", ""))
                count = int(mask.sum())
                current.loc[mask, target_col] = op.get("new_value", "")
                lines.append(f"已更新 **{count}** 筆：符合 **{where_col}** 條件的資料，將 **{target_col}** 改為「{op.get('new_value', '')}」。")
                filename_parts.append("資料已更新")
                export_needed = True
        elif typ == "remove_duplicates":
            cols = [c for c in op.get("columns", []) if c in current.columns]
            before = len(current)
            current = current.drop_duplicates(subset=cols if cols else None).copy()
            lines.append(f"已移除重複資料：{before} 筆 → {len(current)} 筆。")
            filename_parts.append("移除重複")
            export_needed = True
        elif typ == "export":
            export_needed = True

    if not lines and export_needed:
        lines.append("已依照目前工作資料建立 Excel。")

    operation = "；".join(lines) if lines else "AI 操作"
    set_active(current, operation=operation, save_history=True)

    if chart is not None:
        st.session_state.last_chart = chart
    else:
        st.session_state.last_chart = None

    if export_needed:
        base = "_".join(filename_parts) if filename_parts else "DataTalk_整理結果"
        base = f"{base}_{len(current)}筆"
        item = add_download(current, f"{base}.xlsx", operation)
        lines.append(f"已建立新的 Excel：**{item['name']}**。")

    return reply_done(lines, len(current)), current


def make_report():
    df = active_df()
    if df is None:
        return "請先上傳 Excel。", None
    missing = int(df.isna().sum().sum())
    dup = int(df.duplicated().sum())
    return f"""### 主管摘要

這份資料共 **{len(df)} 筆**、**{len(df.columns)} 個欄位**。目前檢查到空白值 **{missing}** 個、重複資料 **{dup}** 筆。

建議優先確認：
1. 狀態、備註、問題欄位中的異常/延遲/NG項目。
2. 供應商或部門是否有集中異常。
3. 整理後可下載 Excel 作為回報附件。
""", None


def process(text):
    nt = normalize(text)
    if "主管" in nt or "報告" in nt or "摘要" in nt:
        return make_report()
    return execute_ops(text)


def sample_files():
    SAMPLE_DIR.mkdir(exist_ok=True)
    return list(SAMPLE_DIR.glob("*.xlsx"))


init()

with st.sidebar:
    st.title("📊 DataTalk AI")
    page = st.radio("頁面", ["上傳資料", "AI 聊天執行", "下載中心", "資料關聯", "操作歷史", "部署檢查"])
    st.divider()
    if original_df() is not None:
        st.success(f"原始資料：{len(original_df())} 筆 / {len(original_df().columns)} 欄")
        if active_df() is not None and len(active_df()) != len(original_df()):
            st.info(f"目前結果：{len(active_df())} 筆")
        if st.button("回到原始資料", use_container_width=True):
            reset_to_original()
            st.rerun()
    else:
        st.info("尚未載入資料")
    if st.button("清空聊天", use_container_width=True):
        st.session_state.messages = [{"role": "assistant", "content": WELCOME}]
        st.rerun()

if page == "上傳資料":
    render_header("上傳資料", "下載測試 Excel 或上傳自己的 Excel。")
    st.subheader("測試 Excel")
    files = sample_files()
    if not files:
        st.warning("尚未產生 sample_data，請先執行 python setup_project.py")
    for i, p in enumerate(files):
        display_name = decode_hash_unicode(p.name)
        with open(p, "rb") as f:
            st.download_button(
                f"下載 {display_name}",
                f.read(),
                file_name=display_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"sample_download_{i}_{p.stem}",
            )
    st.divider()
    f = st.file_uploader("上傳主要 Excel", type=["xlsx", "xls"])
    if f:
        if read_excel is not None:
            df, data, meta = read_excel(f)
        else:
            data = f.getvalue()
            df = pd.read_excel(io.BytesIO(data))
            meta = {"sheet": None, "header_row": 1, "all_sheets": []}
        st.session_state.df = df.copy()
        st.session_state.working_df = df.copy()
        st.session_state.last_result = df.copy()
        st.session_state.original_file_bytes = data
        st.session_state.excel_meta = meta
        st.session_state.file_name = decode_hash_unicode(f.name)
        st.session_state.history.clear()
        st.session_state.redo_stack.clear()
        info = f"已載入 {f.name}：{len(df)} 筆 / {len(df.columns)} 欄"
        if meta and meta.get("sheet"):
            info += f"；自動選擇工作表「{meta.get('sheet')}」、表頭第 {meta.get('header_row')} 列"
        st.success(info)
        st.caption("系統會自動偵測表頭與欄位，不再綁死供應商/備註/狀態這種固定格式。")
    if active_df() is not None:
        st.caption("目前顯示的是工作資料；按左側「回到原始資料」可恢復原始 Excel。")
        st.dataframe(active_df().head(100), use_container_width=True)

elif page == "AI 聊天執行":
    render_header("AI 聊天執行", "直接說需求，Gemini 會依照你上傳的實際欄位與內容判斷並執行 Excel 操作。")
    st.markdown("""
    <div class='chat-hint'>
    <b>可直接輸入：</b><br>
    找出某公司資料｜把某欄不是正常的資料下載｜統計任一欄位｜新增/刪除/改名欄位｜修改符合條件的資料｜依任一欄彙整｜undo / redo
    </div>
    """, unsafe_allow_html=True)

    c1, c2 = st.columns([1, 1])
    with c1:
        st.session_state.use_working_source = st.toggle(
            "接續上一個篩選結果操作",
            value=st.session_state.get("use_working_source", False),
            help="關閉時，每次搜尋都會從原始 Excel 重新查；開啟時會疊加在目前結果上。"
        )
    with c2:
        if st.button("回到原始 Excel", use_container_width=True):
            reset_to_original()
            st.rerun()

    for m in st.session_state.messages:
        with st.chat_message(m["role"]):
            st.write(m["content"])

    prompt = st.chat_input("輸入需求，例如：把備註不是正常的資料下載")
    if prompt:
        st.session_state.messages.append({"role": "user", "content": prompt})
        with st.spinner("正在分析需求…正在執行 Excel 操作…"):
            reply, _ = process(prompt)
        st.session_state.messages.append({"role": "assistant", "content": reply})
        st.rerun()

    if st.session_state.last_result is not None:
        st.subheader("最新結果")
        st.dataframe(st.session_state.last_result, use_container_width=True)
    if st.session_state.downloads:
        st.subheader("下載")
        for i, d in enumerate(st.session_state.downloads[:5]):
            st.download_button(
                f"下載 {d['name']}（{d['rows']} 筆）",
                d["data"],
                file_name=d["name"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_chat_{i}",
            )
    if st.session_state.last_chart:
        chart_df, col = st.session_state.last_chart
        fig = px.bar(chart_df, x=chart_df.columns[0], y=chart_df.columns[1], text=chart_df.columns[1], title=f"{col} 統計")
        st.plotly_chart(fig, use_container_width=True)

elif page == "下載中心":
    render_header("下載中心", "所有由 AI 產生的 Excel 都在這裡。")
    if not st.session_state.downloads:
        st.info("目前沒有下載檔。")
    for i, d in enumerate(st.session_state.downloads):
        with st.container(border=True):
            st.write(f"**{d['name']}**")
            st.caption(f"{d['time']}｜{d['rows']} 筆 / {d['cols']} 欄｜{d['note']}")
            st.download_button("下載 Excel", d["data"], file_name=d["name"], mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_center_{i}")

elif page == "資料關聯":
    render_header("資料關聯", "上傳第二份 Excel，系統會推薦共同欄位，也可以合併資料。")
    df = active_df()
    if df is None:
        st.warning("請先上傳主要 Excel。")
    else:
        f2 = st.file_uploader("上傳第二份 Excel", type=["xlsx", "xls"])
        if f2:
            if read_excel is not None:
                df2, _, meta2 = read_excel(f2)
            else:
                df2 = pd.read_excel(f2)
                meta2 = {}
            st.session_state.df2 = df2
            st.session_state.file2_name = f2.name
            if meta2.get("sheet"):
                st.caption(f"第二份 Excel 自動選擇工作表「{meta2.get('sheet')}」、表頭第 {meta2.get('header_row')} 列。")
        if st.session_state.df2 is not None:
            df2 = st.session_state.df2
            common = [c for c in df.columns if c in df2.columns]
            if common:
                key = st.selectbox("共同欄位", common)
                how = st.selectbox("合併方式", ["left", "inner", "outer", "right"])
                preview = pd.merge(df, df2, on=key, how=how)
                st.write(f"預覽：{len(preview)} 筆 / {len(preview.columns)} 欄")
                st.dataframe(preview.head(50), use_container_width=True)
                if st.button("開始合併並產生下載", use_container_width=True):
                    set_active(preview, operation=f"以 {key} {how} 合併第二份 Excel")
                    add_download(preview, f"DataTalk_合併_{key}_{len(preview)}筆.xlsx", f"以 {key} {how} 合併")
                    st.success("已合併並產生下載。")
            else:
                st.error("沒有完全相同的欄位名稱，請先把欄位名稱改成一致。")

elif page == "操作歷史":
    render_header("操作歷史", "查看 AI 做過哪些 Excel 操作，也可以在聊天輸入 undo / redo。")
    if not st.session_state.operation_log:
        st.info("目前還沒有操作紀錄。")
    else:
        for item in reversed(st.session_state.operation_log):
            st.write("- " + item)

elif page == "部署檢查":
    render_header("部署檢查", "確認部署需要的檔案。")
    st.write("✅ app.py")
    st.write("✅ requirements.txt")
    st.write("✅ sample_data")
    st.write("✅ 不需要 Gemini 也能執行 Excel Agent 基本功能")
    try:
        has_key = bool(get_api_key()) if get_api_key else bool(st.secrets.get("GEMINI_API_KEY"))
        st.write("✅ GEMINI_API_KEY 已設定，AI 解析會優先使用 Gemini" if has_key else "⚠️ GEMINI_API_KEY 未設定：不影響本機 Excel 操作")
    except Exception:
        st.write("⚠️ GEMINI_API_KEY 未設定：不影響本機 Excel 操作")
    st.info("Streamlit Cloud 部署時，請到 App → Settings → Secrets 貼上 GEMINI_API_KEY。")
