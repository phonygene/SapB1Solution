# Session Off - 完整存檔並下班

完整存檔並**結束工作階段**。

## 參數

- 無參數：正常模式，記錄所有工作內容
- `-s` 或 `--selective`：選擇性記錄模式

## 執行步驟

### 正常模式

1. **總結本次對話的所有工作成果**

2. **更新專案狀態**
   - `.claude/shared/project-status.md`
   - `.claude/shared/active-tasks.json`

3. **記錄到 Work Logs**
   - 路徑：`work-logs/daily/YYYY-MM/YYYY-MM-DD.md`
   - 內容：每個任務的完整紀錄

4. **檢查是否月底**
   - 如果是月底，觸發月度彙整
   - 產生 `work-logs/monthly/YYYY-MM.md`

5. **向使用者報告**
   - 已完成任務
   - 未完成任務
   - 下次啟動建議

6. **向使用者道別，結束工作階段**

### 選擇性記錄模式（`-s`）

1. 列出本次會話的所有工作項目
2. 使用 AskUserQuestion 讓使用者選擇排除項目
3. 只記錄保留的內容
4. 其餘同正常模式

## 每日紀錄格式

```markdown
# YYYY-MM-DD 工作紀錄

## {task-id} | {type} | {status}
- **目標**: ...
- **Agent**: backend/ui-ux
- **問題**: ...（如有）
- **解法**: ...（如有）
- **關鍵學習**: ...
- **連結**: commit:xxx, 檔案路徑
- **時間**: HH:MM-HH:MM

---

## 當日摘要
- 完成任務數：N
- 進行中：N
- 主要問題：...
- 更新的 Skills：...
```

## 月底額外動作

如果是當月最後一個工作日：

1. 讀取當月所有 daily logs
2. 產生月度彙整 `work-logs/monthly/YYYY-MM.md`
3. 識別重複問題模式
4. 更新 `work-logs/insights/patterns.md`
5. 提出 skills/ 優化建議
