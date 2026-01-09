# MANAGER Agent - 當前工作

## 狀態
idle

## 當前任務

等候新指令

## 今日完成

1. 假日順延邏輯調整（付款條件不提示、手動改日期詢問）
2. DocumentSearch CSS 變數缺失修復
3. Git commit 審查 + _archive 加入 .gitignore

## 最後更新
2026-01-08 - commit 1356c68: holiday prompt + CSS fix

## 已完成項目

- Git Worktree 設置
  - `/c/Projects/worktrees/backend` - agent/backend 分支
  - `/c/Projects/worktrees/ui-ux` - agent/ui-ux 分支
  - 三個工作目錄已同步到 commit 4fb1ff8

- 假日跳過到期日功能
  - 建立 jHolidays 資料表
  - 建立 HolidayHelper 共用模組
  - 費用申請單到期日自動跳過假日並顯示提示
  - 請購單需求日自動跳過假日並顯示提示
  - 匯入 2026 年台灣政府行事曆（21 筆）
- 供應商價格更新邏輯
  - 選擇供應商時，若有項目則詢問是否更新價格
  - 從 OPOR/POR1 查詢供應商最後採購價
  - 若無記錄則使用 OITM.LastPurPrc
- 倉庫查詢修正
  - SQL 改為 `SELECT T0.[WhsCode], T0.[WhsName] FROM OWHS T0`
  - 移除 Inactive 條件
- 審核權限欄位 (PU_App)
  - User 表新增 PU_App 欄位
  - CheckApprovalPermission 改用 PU_App
- 工具錯誤記錄機制
- 請購單邏輯修正（參照費用申請單）
  - 稅碼改為硬編碼：1-應稅(5%), 2-零稅(0%), 3-免稅(0%)
  - 新增含稅單價 (PriceAfVAT) 欄位
  - 稅額可手動編輯
  - 金額計算邏輯改為 Math.Floor() 無條件捨去
  - 產品別/部門別查詢改為排除 Centr%
- DDL 已執行：jOPRQ, jPRQ1 資料表已存在
- Littlebird 輸入鎖定功能
- UTF-8 BOM 編碼問題修正

## 進行中的變更

（無）

## 待處理項目

- Phase 2：轉採購訂單功能
- Phase 2：PDF 匯出功能
