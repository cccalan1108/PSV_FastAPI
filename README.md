# 使用說明

## 專案部署說明

1. 安裝 uv（如尚未安裝）：
   ```bash
   pip install uv
   ```
2. 安裝專案依賴：
   ```bash
   uv pip install -r requirements.txt
   ```
3. 啟動 API 伺服器：
   ```bash
   uvicorn main:app --reload
   ```
   或用 uv 直接執行（推薦）：
   ```bash
   uvicorn main:app --host 0.0.0.0 --port 8000
   ```
   預設網址為 http://127.0.0.1:8000

4. 注意事項：
   - 請確保 `pyproject.toml` 與 `Data Sheet.xlsm`、`Calculation Sheet.xlsm` 已放在專案根目錄。
   - 若有權限問題，請用管理員權限執行。
   - 若需在區網其他電腦存取，請用 `--host 0.0.0.0`。

## curl 測試範例

### Data Sheet 轉 Calculation Sheet
```bash
curl -X 'POST' \
  'http://127.0.0.1:8000/data2calc/' \
  -H 'accept: application/json' \
  -H 'Content-Type: multipart/form-data' \
  -F 'data_sheet=@Data Sheet.xlsm;type=application/vnd.ms-excel.sheet.macroEnabled.12'
```

### Calculation Sheet 轉 Data Sheet
```bash
curl -X 'POST' \
  'http://127.0.0.1:8000/calc2data/' \
  -H 'accept: application/json' \
  -H 'Content-Type: multipart/form-data' \
  -F 'calc_sheet=@Calculation Sheet.xlsm;type=application/vnd.ms-excel.sheet.macroEnabled.12'
```

## 常見錯誤排除

- **422 Unprocessable Entity**：
  - 請確認有正確上傳檔案，欄位名稱需為 `data_sheet` 或 `calc_sheet`。
- **Conversion failed**：
  - 請確認上傳的 Excel 檔案內容正確，且格式符合範本要求。
- **找不到檔案或權限錯誤**：
  - 請確認範本檔案（如 Data Sheet.xlsm）已放在伺服器端正確路徑，且有讀寫權限。
- **API 無法啟動**：
  - 請確認所有必要套件已安裝，且 Python 版本相容。