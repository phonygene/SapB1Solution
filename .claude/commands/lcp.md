---
description: 記錄 Log、Commit 並 Push (project)
---

# Log & Commit & Push

記錄工作日誌、執行 git commit，並 push 到遠端。

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

### 5. Push 到遠端

```bash
git push
```

若有上游分支衝突，先提示用戶確認再處理。

### 6. 回報結果

顯示 commit 與 push 結果。

## 注意事項

- 遵循專案的 commit message 規範
- Push 前會顯示即將推送的 commits
- 若只要 commit 不 push，請使用 `/lc`
