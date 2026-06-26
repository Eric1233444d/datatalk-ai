import re
import pandas as pd
from services.excel_service import find_column

ABNORMAL_PATTERN = "異常|延遲|失敗|NG|ng|問題|缺失|未完成|待處理|風險|錯誤"

def summarize_df(df):
    return f"""這份資料共有 {len(df)} 筆、{len(df.columns)} 個欄位。
空白值：{int(df.isna().sum().sum())}
重複資料：{int(df.duplicated().sum())}
欄位：{", ".join(map(str, df.columns))}"""

def count_column(df, col):
    out = df[col].fillna("空白").astype(str).value_counts().reset_index()
    out.columns = [col, "數量"]
    return out

def search_all(df, keyword):
    mask = df.astype(str).apply(lambda row: row.str.contains(keyword, case=False, na=False).any(), axis=1)
    return df[mask]

def detect_abnormal(df):
    candidate_cols = [c for c in df.columns if any(k in str(c) for k in ["狀態","結果","備註","問題","異常","說明"])]
    if not candidate_cols:
        return pd.DataFrame(), "找不到可判斷異常的欄位。"
    mask = pd.Series(False, index=df.index)
    for c in candidate_cols:
        mask = mask | df[c].astype(str).str.contains(ABNORMAL_PATTERN, case=False, na=False)
    result = df[mask]
    return result, f"已找到 {len(result)} 筆可能異常資料。"

def filter_condition(df, text):
    col = find_column(text, df.columns)
    if col is None:
        return None, "找不到要篩選的欄位。"
    m = re.search(r"(是|為|等於|包含)\s*(.+)", text)
    if not m:
        return None, f"請指定條件，例如：篩選 {col} 是 異常。"
    value = m.group(2).strip()
    result = df[df[col].astype(str).str.contains(value, case=False, na=False)]
    return result, f"已篩選「{col}」包含「{value}」的資料，共 {len(result)} 筆。"

def replace_value(df, text):
    # 把 未完成 取代成 待處理 / 將 未完成 改成 待處理
    m = re.search(r"(把|將)\s*(.+?)\s*(取代成|改成|改為)\s*(.+)", text)
    if not m:
        return df, "請用：把 未完成 取代成 待處理。"
    old, new = m.group(2).strip(), m.group(4).strip()
    new_df = df.copy()
    count = 0
    for col in new_df.columns:
        s = new_df[col].astype(str)
        hit = s == old
        count += int(hit.sum())
        new_df.loc[hit, col] = new
    return new_df, f"已將所有完全符合「{old}」的值改成「{new}」，共修改 {count} 格。"

def rename_column(df, text):
    m = re.search(r"(欄位)?\s*(把|將|重新命名)?\s*(.+?)\s*(改成|改為|為)\s*(.+)", text)
    if not m:
        return df, "請用：把 狀態 改成 處理狀態。"
    old_raw, new_col = m.group(3).strip(), m.group(5).strip()
    old_col = find_column(old_raw, df.columns)
    if old_col is None:
        return df, f"找不到欄位「{old_raw}」。"
    return df.rename(columns={old_col: new_col}), f"已將欄位「{old_col}」改成「{new_col}」。"

def fill_blank(df, text):
    m = re.search(r"(補成|填成|補為|填為)\s*(.+)", text)
    value = m.group(2).strip() if m else "0"
    new_df = df.copy().fillna(value)
    return new_df, f"已將空白值補成「{value}」。"

def add_column(df, text):
    # 新增欄位 風險等級 值為 待確認
    m = re.search(r"新增欄位\s*(.+?)(\s*值為\s*(.+))?$", text)
    if not m:
        return df, "請用：新增欄位 風險等級 值為 待確認。"
    col = m.group(1).strip()
    value = m.group(3).strip() if m.group(3) else ""
    new_df = df.copy()
    new_df[col] = value
    return new_df, f"已新增欄位「{col}」，預設值為「{value}」。"

def drop_column(df, text):
    col = find_column(text, df.columns)
    if col is None:
        return df, "找不到要刪除的欄位。"
    return df.drop(columns=[col]), f"已刪除欄位「{col}」。"

def sort_df(df, text):
    col = find_column(text, df.columns)
    if col is None:
        return df, "找不到排序欄位。"
    ascending = not ("由大到小" in text or "降冪" in text or "最高" in text)
    return df.sort_values(by=col, ascending=ascending), f"已依照「{col}」{'由小到大' if ascending else '由大到小'}排序。"

def merge_dfs(df, df2, key_col, how):
    return pd.merge(df, df2, on=key_col, how=how)
