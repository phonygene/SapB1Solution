# Session On - 上班/開始工作

開始工作階段，載入專案狀態。

## 執行步驟

1. **讀取專案狀態**
   - `.claude/shared/project-status.md`
   - `.claude/shared/active-tasks.json`

2. **檢查待處理任務**
   - 未完成的任務
   - blocked 狀態的任務

3. **讀取最近工作紀錄**
   - `work-logs/daily/` 中最近的紀錄

4. **向使用者報告**
   - 專案整體狀態
   - 未完成任務
   - 今日優先事項
   - 下一步建議

## Agent 特定初始化

### Manager Session
```bash
# 啟動檔案監聽（監控 Agent 完成任務）
inotifywait -m -r -e close_write .claude/handoff/ | while read FILE; do
    [[ "$FILE" == *"output.md" ]] && echo "🔔 Task 完成！"
done &
```

### Backend Session
- 讀取 `.claude/agents/BACKEND.md`
- 讀取 `skills/backend-checklist.md`
- 讀取 `skills/sap-checklist.md`

### UI-UX Session
- 讀取 `.claude/agents/UI-UX.md`
- 讀取 `skills/ui-checklist.md`
- 讀取 `skills/ui-design-system.md`

## 報告格式

```markdown
# Session Started

## 專案狀態
- 版本：vX.Y.Z
- 分支：{branch}

## 未完成任務
| ID | 任務 | 狀態 | Agent |
|----|------|------|-------|
| ... | ... | ... | ... |

## 今日優先
1. ...
2. ...

## 建議下一步
...
```
