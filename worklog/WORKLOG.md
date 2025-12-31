# 工作記錄

> 即時記錄每次請求與結果，取代 Session 彙整機制

---

## 2024-12-31

### 10:30 - DI API 稅額寫入修正
- 要求：寫入用戶輸入的稅額到 SAP，避免被 SAP 重算
- 結果：新增 Lines.TaxTotal 寫入邏輯
- 檔案：ExpenseClaimForm.aspx.vb
- Commit：v1.0.8

---

## 2024-12-30

### 18:00 - Commit v1.0.7
- 要求：commit 網站美化與工作流程規範
- 結果：已 commit，版號 1.0.6 → 1.0.7
- 檔案：15 files changed

### 17:30 - 建立 worklog 記錄結構
- 要求：建立即時記錄機制取代 sess 管理
- 結果：創建 WORKLOG.md, DECISIONS.md, ISSUES.md
- 檔案：worklog/

### 17:25 - 更新工作流程規範
- 要求：SQL 用 shopfloor 流程，程式碼改用 Git
- 結果：已更新 workflow-standards.md
- 檔案：agent-os/standards/global/workflow-standards.md

### 17:00 - 美化網站介面
- 要求：Index.aspx 首頁只顯示登入按鈕，登入後進入功能選單
- 結果：創建 Home.aspx，修改 MySite1.Master 美化，修改 login 導向
- 檔案：Index.aspx, Home.aspx, MySite1.Master, login.aspx.vb
- 問題：Index.aspx 中文亂碼（缺 BOM）、Home.aspx 未加入專案
- 修正：已加 BOM，已更新 .vbproj

---

## 記錄格式

```
### HH:MM - 簡短標題
- 要求：用戶要求的內容
- 結果：執行結果摘要
- 檔案：涉及的檔案
- 問題：（如有）遇到的問題
- 待辦：（如有）後續待處理
```
