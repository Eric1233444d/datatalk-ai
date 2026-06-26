# DataTalk AI v2

這版已升級成 Excel Agent：使用者用自然語言輸入，系統會直接執行篩選、搜尋、統計、排序、匯出 Excel，不只是回答建議。

## 執行方式

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 可直接輸入的範例

- 把供應商是華新工程的資料建一個 Excel
- 把備註不是正常的資料下載下來
- 搜尋異常
- 統計狀態
- 把華新工程抓出來，再依日期排序，匯出 Excel
- undo
- redo

## 這版新增功能

- 自然語言轉 Excel 操作
- 對話式多步操作
- 否定條件理解，例如「備註不是正常」
- 自動命名下載 Excel
- 拼字容錯，例如「華興工程」會嘗試比對成資料中的「華新工程」
- Undo / Redo
- 操作歷史頁面
- 執行動畫與更自然的聊天回覆
- Gemini API Key 非必填，沒設定也能操作 Excel

## 注意

`.streamlit/secrets.toml` 已清空，請自行填入 Gemini Key。不要把自己的 API Key 上傳到公開 GitHub。

## Gemini Key 部署方式

本機測試可以使用 `.streamlit/secrets.toml`：

```toml
GEMINI_API_KEY = "你的 Gemini API Key"
```

Streamlit Cloud 部署時，請到：

`App → Settings → Secrets`

貼上同樣的 TOML 設定後重新 Deploy。

程式會優先使用 Gemini 解析自然語言；如果 Gemini Key 無效或額度不足，仍會自動退回本機 Excel Agent 規則，不會整個壞掉。
