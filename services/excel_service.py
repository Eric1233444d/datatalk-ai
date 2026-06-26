import io
import re
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

SAMPLE_DIR = Path(__file__).resolve().parents[1] / "sample_data"

def read_excel(uploaded_file):
    data = uploaded_file.getvalue()
    df = pd.read_excel(io.BytesIO(data))
    return df, data

def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="DataTalk_Result")
        ws = writer.book["DataTalk_Result"]
        header_fill = PatternFill("solid", fgColor="1F2937")
        header_font = Font(color="FFFFFF", bold=True)
        thin = Side(style="thin", color="CBD5E1")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")
            cell.border = border
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(horizontal="center")
        for col in ws.columns:
            letter = col[0].column_letter
            max_len = max(len(str(c.value)) if c.value is not None else 0 for c in col)
            ws.column_dimensions[letter].width = min(max_len + 4, 35)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    return output.getvalue()

def preserve_style_export(original_bytes: bytes | None, df: pd.DataFrame) -> bytes:
    """
    優先保留原始 workbook 第一張表的樣式：只覆寫既有格子的 value。
    若資料尺寸比原本大，新增的格子會複製同欄上一列樣式。
    如果沒有原始檔，就輸出美化後 Excel。
    """
    if not original_bytes:
        return dataframe_to_excel_bytes(df)

    try:
        wb = load_workbook(io.BytesIO(original_bytes))
        ws = wb.active
        max_row_needed = len(df) + 1
        max_col_needed = len(df.columns)

        # header
        for c_idx, col_name in enumerate(df.columns, start=1):
            ws.cell(row=1, column=c_idx).value = col_name

        # data
        for r_idx, row in enumerate(df.itertuples(index=False), start=2):
            for c_idx, value in enumerate(row, start=1):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.value = None if pd.isna(value) else value
                # copy style from previous row when new row is outside original styled area
                if r_idx > ws.max_row and r_idx > 2:
                    src = ws.cell(row=r_idx-1, column=c_idx)
                    cell._style = src._style

        # clear extra old rows/cols values, keep style
        for r in range(max_row_needed + 1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                ws.cell(r, c).value = None

        if ws.max_column > max_col_needed:
            for c in range(max_col_needed + 1, ws.max_column + 1):
                for r in range(1, ws.max_row + 1):
                    ws.cell(r, c).value = None

        output = io.BytesIO()
        wb.save(output)
        return output.getvalue()
    except Exception:
        return dataframe_to_excel_bytes(df)

def get_sample_files():
    return sorted(SAMPLE_DIR.glob("*.xlsx"))

def find_column(text, columns):
    text_lower = str(text).lower()
    for col in columns:
        if str(col).lower() in text_lower:
            return col
    aliases = {
        "狀態": ["狀態", "結果", "status"],
        "部門": ["部門", "單位", "department"],
        "供應商": ["供應商", "廠商", "vendor", "supplier"],
        "負責人": ["負責人", "承辦", "owner"],
        "日期": ["日期", "時間", "date"],
        "區域": ["區域", "廠區", "area"],
        "類別": ["類別", "分類", "category"],
        "金額": ["金額", "費用", "amount", "price", "cost"],
        "數量": ["數量", "qty", "quantity", "count"],
        "工單": ["工單", "工單編號", "id", "編號"],
    }
    for _, keys in aliases.items():
        if any(k.lower() in text_lower for k in keys):
            for col in columns:
                if any(k.lower() in str(col).lower() for k in keys):
                    return col
    return None
