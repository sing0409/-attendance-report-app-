# 出勤報表自動彙整系統

一個基於 Streamlit 的 Web 應用程式，可自動彙整員工出勤紀錄並生成符合格式的月報表。

## 功能特性

- 📊 **自動彙整出勤紀錄** - 從 Excel 檔案讀取出勤資料
- 📝 **智能時數計算** - 全天休假自動計為 8 小時（已扣除中午休息）
- 🏢 **假別分類** - 自動分類年休、補休、病假、事假、生理假等
- ⏰ **加班原因提取** - 從外出登記自動提取加班事由
- 📋 **Word 格式輸出** - 直接填入 Word 範本，格式規範
- ⚡ **快速處理** - 5-10 秒內完成彙整

## 使用方式

1. 打開網站
2. 上傳當月的 **Excel 出勤紀錄**
3. 上傳您的 **Word 月報表範本**
4. 點擊「開始彙整」
5. 下載完成的 Word 報表

## 系統需求

- Python 3.8+
- 依賴套件見 `requirements.txt`

## 本地運行

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 部署到 Streamlit Cloud

1. 將此倉庫推送到 GitHub
2. 在 [Streamlit Cloud](https://streamlit.io/cloud) 上連接您的 GitHub 倉庫
3. 選擇此檔案作為主應用程式
4. 部署完成！

## 技術棧

- **Streamlit** - Web 框架
- **Pandas** - 資料處理
- **python-docx** - Word 檔案處理
- **openpyxl** - Excel 檔案讀取

## 授權

MIT License

## 作者

由 Manus AI 開發
