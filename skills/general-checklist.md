# 通用開發檢查清單

> 所有 Agent 共用的檢查項目

---

## Git 提交規範

- [ ] Commit message 包含 task-id
- [ ] 不要 commit 未完成的代碼
- [ ] 不要 commit 含有密碼/金鑰的檔案

## 檔案編碼

- [ ] 所有檔案使用 **UTF-8 with BOM**
- [ ] 中文內容不會變成亂碼

## ASP.NET 控制項

- [ ] 新增控制項時同步更新 `.aspx.designer.vb`
- [ ] `<%@ Page ... Inherits="MgmSP.ClassName" %>` 必須有 `MgmSP.` 前綴

## 程式碼品質

- [ ] 有適當的錯誤處理（Try-Catch）
- [ ] 敏感操作有記錄到 EventLog
- [ ] 不使用 magic number，用常數定義

---

## 工作流程規範（絕對遵守）

> **這不是建議，是強制規定。任何技術工作開始前必須先確認流程。**
> **違反流程造成的損失無法估量，絕對不能因為其他目的著急而跳過。**

### 每次工作開始前（必做）
1. 確認今日 worklog 存在：`work-logs/daily/YYYY-MM/YYYY-MM-DD.md`
2. 確認沒有未 push 的 commits：`git status`
3. 如果有未完成的流程，先完成流程再開始新工作

### 每次修改後（必做）
1. 立即 commit（commit message 包含 task-id）
2. 更新 worklog 記錄該修改
3. Push 到遠端

### 每次工作結束前（必做）
1. 確認所有修改已 commit 並 push
2. 確認 worklog 已更新
3. 如果有新學習，更新 skills/*.md

### 違反的後果
- 未 commit/push：崩潰時工作遺失，需要重做
- 未記錄 worklog：無法追蹤歷史，浪費未來的時間
- 未更新 skills：同樣的錯誤會重複發生

---

## 從錯誤中學習（持續新增）

> 每次犯錯就新增一條

### 2026-01-05: 工作流程疏忽
- **問題**: 崩潰恢復後只專注技術修復，忽略 commit/worklog/skills 更新
- **根因**: 將「解決問題」置於「遵守流程」之上
- **教訓**: 工作流程是絕對優先的，不是「先做完再說」
- **改進**: 在任何技術工作前，先確認流程狀態

### 2026-01-05: Windows API 驗證
- **問題**: `GetWindowThreadProcessId` 用法錯誤（混淆返回值與 output parameter）
- **教訓**: Windows API 操作必須驗證結果，閱讀文檔確認參數用法
- **正確用法**: `thread_id = GetWindowThreadProcessId(hwnd, None)` - 返回值是 thread ID
