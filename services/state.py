import streamlit as st

def init_state():
    defaults = {
        "df": None,
        "df2": None,
        "working_df": None,
        "file_name": None,
        "file2_name": None,
        "original_file_bytes": None,
        "result_df": None,
        "messages": [
            {"role": "assistant", "content": "請先到「上傳資料」下載測試 Excel 或上傳自己的 Excel。之後你可以直接問我：搜尋異常、統計狀態、哪個部門異常最多、把空白值補成 0、把狀態改成處理狀態、產生主管報告。"}
        ],
        "last_chart_df": None,
        "last_chart_title": None,
        "operation_history": [],
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

def active_df():
    return st.session_state.working_df

def set_active_df(df, operation=None):
    st.session_state.working_df = df
    st.session_state.result_df = df
    if operation:
        st.session_state.operation_history.append(operation)
