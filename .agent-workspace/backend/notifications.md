# Backend Agent 通知

---

## 2026-01-09 15:55 - B3-B4 任務確認

**任務ID**：2026-01-08-user-required-fields-backend
**優先級**：High
**狀態**：請先自我檢查，再執行未完成項目

### 背景說明

之前因 Worktree 同步問題，你讀取到的是舊規格 `user-profile-backend`，而非新規格 `user-required-fields-backend`。現已透過 symlink 修復同步問題。

### 請先自我檢查

1. 讀取正確的規格：`.agent-workspace/handoff/2026-01-08-user-required-fields-backend/spec.md`
2. 檢查 `MgmSP/ExpenseClaimForm.aspx.vb` 中：
   - `CheckUserExpDept()` 是否已包含 EmpSeries 檢查？
   - `btnExpDeptConfirm_Click` 是否已儲存 EmpSeries？
3. 如果已完成，請回報；如果未完成，請繼續執行 B3-B4

### 待確認/執行項目

| 編號 | 任務 | 預期狀態 |
|------|------|----------|
| B1 | DDL: EmpSeries 欄位 | ✓ 已完成 |
| B2 | UserProfileHelper.vb | ✓ 已完成 |
| B3 | 修改 CheckUserExpDept | ? 請確認 |
| B4 | 修改 btnExpDeptConfirm_Click | ? 請確認 |

### 規格位置

`.agent-workspace/handoff/2026-01-08-user-required-fields-backend/spec.md`

---

## 2026-01-09 14:30 - 任務啟動通知（舊）

**任務ID**：2026-01-08-user-required-fields-backend
**優先級**：High
**狀態**：請立即開始執行

### 任務內容 (B1-B4)

| 編號 | 任務 | 說明 |
|------|------|------|
| B1 | DDL | `ALTER TABLE [User] ADD EmpSeries NVARCHAR(50) NULL` |
| B2 | 建立 UserProfileHelper.vb | 使用者資料讀寫共用模組 |
| B3 | 修改 CheckUserExpDept | 檢查費用部門 + 工號兩者都有值 |
| B4 | 修改 btnExpDeptConfirm_Click | 儲存時同時存費用部門 + 工號 |

### 規格位置

`.agent-workspace/handoff/2026-01-08-user-required-fields-backend/spec.md`

### 注意事項

- UI-UX 的 U5 任務依賴此任務完成
- 完成後請更新 output.md 並通知 Manager

---

## 2026-01-08 10:30 - 新任務分派（已併入上方）

**任務ID**：2026-01-08-user-profile-backend
**優先級**：High
**標題**：User 表欄位擴充與資料存取模組

### 任務摘要

1. 在 `[User]` 表新增 `EmpSeries`（工號）欄位
2. 建立共用模組 `UserProfileHelper.vb`

### 規格位置

`.agent-workspace/handoff/2026-01-08-user-profile-backend/spec.md`

### 注意事項

- UI-UX Agent 的任務依賴此任務完成
- 請優先處理

---

## 2026-01-06 14:38:00 [ROUND-4]
*jdi* T010-BE。請輸出「T010-BE OK」並寫入 output.md。
