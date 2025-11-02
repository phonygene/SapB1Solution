# Session Off - 完整存檔並下班

完整存檔並**結束工作階段**。

## 執行步驟

1. 總結本次對話的所有工作成果
2. 建立歷史記錄檔案：
   - 路徑：`worklog/checkPoint_YYYYMMDD_HHMM.log`
   - 內容：詳細的工作記錄，包含時間戳記與完整細節
3. 更新最新狀態檔案：
   - 路徑：`worklog/LastCheckPoint.log`
   - 內容：
     - 專案整體狀態
     - 已完成的任務（詳細列表）
     - 未完成的待辦事項（按優先順序）
     - 重要決策記錄
     - 已產生的檔案清單
     - Git 狀態
     - 下次啟動建議
     - 測試前準備檢查清單
4. 向使用者報告存檔完成，並說明下次啟動方式
5. **向使用者道別，結束工作階段**

詳細的檔案格式請參考 `agent-os/standards/global/session-management.md`。
