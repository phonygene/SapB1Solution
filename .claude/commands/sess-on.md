# Session On - 上班/開始工作

讀取 `agent-os/SESSION_INIT.md` 並執行完整的初始化流程。

## 執行步驟

SESSION_INIT.md 會指導你完成以下步驟：

1. 讀取核心規範檔案
2. 讀取專案狀態（worklog/LastCheckPoint.log、TODO.md）
3. 向使用者報告上次工作狀態
4. 等待使用者回應並準備接續工作

請嚴格按照 SESSION_INIT.md 中定義的流程執行。

## ⚠️ 重要提醒

執行完初始化流程後，向使用者提醒：

**本專案採用 Shopfloor 協作模式**：
- 所有程式碼產出必須先輸出到 `shopfloor/Claude_TMP/`
- 不在對話中貼大段程式碼（超過 20 行）
- 等待使用者確認後才加入正式專案
- 詳細規範請參考：`shopfloor/Claude_TMP/etc/README_協作模式說明.txt`
