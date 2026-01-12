---
description: 記錄 Log 並 Commit (project)
---

# Log & Commit

記錄工作日誌並執行 git commit。

## 執行步驟

### 1. 更新工作日誌

檢查並更新 `work-logs/daily/YYYY-MM/YYYY-MM-DD.md`：
- 若檔案不存在，建立新檔案
- 補充今日尚未記錄的工作項目
- 使用標準格式（見 CLAUDE.md）

### 2. 同步待辦事項

若有新增待辦，同步更新 `work-logs/TODO.md`

### 3. 檢查 VERSION

根據本次變更類型判斷是否需要更新版號：
- `feat`: Minor 版號 +1
- `fix/chore/docs`: Patch 版號 +1

### 4. 執行 Git 操作

```bash
git add -A
git status
```

確認變更內容後，產生 commit message 並執行：
```bash
git commit -m "commit message"
```

### 5. 回報結果

顯示 commit 結果與變更摘要。

## 注意事項

- 遵循專案的 commit message 規範
- 不會自動 push，僅 commit 到本地
- 若要 push，請使用 `/lcp`
