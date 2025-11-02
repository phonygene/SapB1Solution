# Session Wrap - 階段性存檔（繼續工作）

階段性存檔，但**不結束對話**。

## 執行步驟

1. 總結本次階段的工作成果
2. 建立歷史記錄檔案：
   - 路徑：`worklog/checkPoint_YYYYMMDD_HHMM.log`
   - 內容：詳細的工作記錄，包含時間戳記與完整細節
3. 更新最新狀態檔案：
   - 路徑：`worklog/LastCheckPoint.log`
   - 內容：最新的專案狀態快照、待辦清單、下次啟動建議
4. 向使用者提供彙整報告
5. **繼續工作**，不結束對話

詳細的檔案格式請參考 `agent-os/standards/global/session-management.md`。
