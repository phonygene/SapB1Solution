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

## 從錯誤中學習（持續新增）

> 每次犯錯就新增一條

*（待累積）*
