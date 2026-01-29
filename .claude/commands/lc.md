---
description: 記錄 Log 並 Commit (project)
allowed-tools: Bash, Write, Read, Glob
argument-hint: [commit-message]
---

# Log & Commit (LC)

你現在必須執行標準的提交流程。

## 當前環境資訊
**日期**: !`date +"%Y-%m-%d"`
**Git 狀態**:
!`git status -s`

## 用戶參數
**Commit Message**: $ARGUMENTS

## 執行流程（必須嚴格遵守）

1.  **更新工作日誌** (必要!)
    - 讀取 `work-logs/daily/` 下今日的日誌（若無則新建）。
    - 將上述 Git 變更摘要寫入日誌。
    - **禁止**在未更新日誌的情況下 Commit。

2.  **執行 Commit**
    - 使用 `bash` 工具。
    - 指令：`git add .`
    - 指令：`git commit -m "..."`
    - 若用戶有提供 `$ARGUMENTS`，直接使用作為 message。
    - 若無，則根據 `git status` 內容自動撰寫符合 Conventional Commits 的訊息。

3.  **檢查版號** (可選)
    - 若變更涉及新功能 (`feat`)，檢查 `VERSION` 檔案並建議是否升級 Minor 版號。

4.  **最終回報**
    - 顯示 Commit Hash 和 Message。
    - 提醒用戶若要推送可使用 `/lcp`。
