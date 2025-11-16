# Session 管理規範

**版本**：2.0
**更新日期**：2025-11-14

---

## 概述

Session 管理機制讓你能夠：
- 追蹤工作進度
- 在工作中斷後快速恢復
- 保持專案狀態的連續性
- 明確的開始與結束工作流程

---

## 四個核心指令

### 1. sess-on（上班/開始工作）

**用途**：開始新的工作 Session，讀取上次工作狀態

**執行方式**：
- Slash Command: `/sess-on`
- 純文字: `Claude, sess on.`

**執行流程**：
1. 讀取核心規範檔案（`SESSION_INIT.md` 中定義）
2. 讀取工作狀態（`worklog/LastCheckPoint.log`）
3. 讀取待辦事項（`TODO.md`）
4. 向使用者報告：
   - 上次工作時間
   - 當前專案狀態
   - 未完成的待辦事項
   - 建議的下一步
5. 提醒協作模式
6. 等待使用者指示

**範例輸出**：
```markdown
📋 Session 初始化完成

專案：MyAPI (FastAPI 專案)
上次工作時間：2025-11-13 18:30 UTC

當前進度：
✅ 使用者資料模型 - 已完成
✅ 資料庫遷移 - 已完成
⏳ 使用者認證 API - 進行中（70%）

未完成的待辦事項：
1. [高] 完成 JWT 認證功能
2. [中] 撰寫單元測試
3. [低] 更新 API 文件

⚠️ 協作模式提醒：
本專案採用 Shopfloor 協作模式
- 所有程式碼會先輸出到 shopfloor/Claude_TMP/
- 請在 IDE 中檢視並執行
- 簡短回報結果即可

建議下一步：繼續完成 JWT 認證功能

請問要從哪裡開始？
```

---

### 2. sess-check（檢查進度）

**用途**：快速查看當前工作進度，不寫入任何檔案

**執行方式**：
- Slash Command: `/sess-check`
- 純文字: `Claude, sess check.`

**執行流程**：
1. 讀取 `worklog/LastCheckPoint.log`
2. 讀取 `TODO.md`
3. 報告當前狀態
4. **不寫入任何檔案**

**範例輸出**：
```markdown
📊 當前進度檢查

專案狀態：
- 當前階段：Sprint 1 - 使用者認證
- 完成度：70%
- 上次更新：2025-11-13 18:30 UTC

進行中的任務：
⏳ JWT 認證功能（70%）

待辦事項：
🔥 高優先級（1 項）
⚡ 中優先級（3 項）
💡 低優先級（2 項）

建議：繼續完成 JWT 認證功能
```

**使用場景**：
- 工作中快速確認進度
- 向團隊成員報告狀態
- 決定下一步工作方向

---

### 3. sess-wrap（階段存檔）

**用途**：階段性存檔當前工作，但繼續 Session

**執行方式**：
- Slash Command: `/sess-wrap`
- 純文字: `Claude, sess wrap.`

**執行流程**：
1. 收集當前工作狀態
2. 更新 `worklog/LastCheckPoint.log`
3. 更新 `TODO.md` 中的完成狀態
4. 報告存檔內容
5. Session 繼續，可以接著工作

**範例輸出**：
```markdown
📝 階段存檔完成

本次完成：
✅ JWT token 生成功能
✅ Token 驗證中介軟體
✅ 單元測試（基本測試）

進行中：
⏳ Token 刷新機制（50%）

已更新：
- worklog/LastCheckPoint.log
- TODO.md

下次繼續：完成 Token 刷新機制

你可以：
- 繼續工作
- 或稍後使用 /sess-on 恢復
```

**使用場景**：
- 午休前快速存檔
- 完成一個小階段
- 切換到其他任務前
- 定期備份工作狀態

---

### 4. sess-off（下班/結束工作）

**用途**：完整存檔並結束 Session

**執行方式**：
- Slash Command: `/sess-off`
- 純文字: `Claude, sess off.`

**執行流程**：
1. 收集完整工作狀態
2. 更新 `worklog/LastCheckPoint.log`（詳細記錄）
3. 更新 `TODO.md`
4. 記錄所有未完成的工作
5. 提供下次開始的建議
6. 報告本次 Session 完整總結

**範例輸出**：
```markdown
📦 完整存檔完成

今日工作總結：
✅ 完成項目（3 項）
  - JWT token 生成功能
  - Token 驗證中介軟體
  - 單元測試（基本測試）

⏳ 進行中（1 項）
  - Token 刷新機制（50%）

📋 待辦事項狀態
  - 高優先級：1 項完成，1 項進行中
  - 中優先級：3 項待處理
  - 低優先級：2 項待處理

⏰ 工作時間：2025-11-13 14:00 - 18:30 UTC（4.5 小時）

📝 下次建議：
1. 繼續完成 Token 刷新機制
2. 測試完整的認證流程
3. 開始撰寫 API 文件

已更新檔案：
- worklog/LastCheckPoint.log
- TODO.md

明天見！👋
```

**使用場景**：
- 每天下班前
- 完成一個大階段
- 切換到其他專案前
- 需要完整總結時

---

## 檔案規範

### worklog/LastCheckPoint.log

**格式**：
```markdown
# 工作檢查點日誌

**最後更新**：[ISO 8601 格式日期時間]
**Session 狀態**：[進行中 / 已結束]

---

## 當前專案狀態

**專案名稱**：[專案名稱]
**當前階段**：[開發階段]
**完成度**：[百分比]

---

## 已完成項目

✅ [日期時間] - [項目描述]
✅ [日期時間] - [項目描述]

---

## 進行中項目

⏳ [項目描述]
   - 狀態：[百分比] 完成
   - 下一步：[下一步動作]

---

## 待辦事項

### 高優先級
- [ ] [待辦事項]
- [ ] [待辦事項]

### 中優先級
- [ ] [待辦事項]

### 低優先級
- [ ] [待辦事項]

---

## 等待使用者執行的任務

[如果有等待使用者執行的任務，在此列出]

---

## 技術備註

[任何重要的技術筆記、決策記錄]

---

## 下次開始建議

1. [建議1]
2. [建議2]
3. [建議3]
```

---

### TODO.md

**格式**：
```markdown
# 專案待辦事項

**專案名稱**：[專案名稱]
**最後更新**：[日期時間]

---

## 🔥 高優先級

- [ ] [待辦事項1]
  - [x] [子任務1]
  - [ ] [子任務2]

- [ ] [待辦事項2]

---

## ⚡ 中優先級

- [ ] [待辦事項3]
- [ ] [待辦事項4]

---

## 💡 低優先級

- [ ] [待辦事項5]

---

## ✅ 已完成

- [x] [已完成項目1] - [完成日期]
- [x] [已完成項目2] - [完成日期]

---

## 📝 保留功能（暫不實作）

- [未來可能實作的功能]
```

---

## 最佳實踐

### 1. 每日工作流程

**早上**：
```
/sess-on
[查看昨天進度，規劃今天工作]
```

**午休前**：
```
/sess-wrap
[存檔上午工作]
```

**下午**：
```
/sess-check
[確認進度]
```

**下班前**：
```
/sess-off
[完整存檔，總結今日工作]
```

### 2. 長期專案

**每週回顧**：
- 使用 `sess-check` 查看整週進度
- 更新長期待辦事項
- 調整優先級

**Sprint 切換**：
- 使用 `sess-off` 完整記錄當前 Sprint
- 更新 `SESSION_INIT.md` 中的「當前階段」
- 使用 `sess-on` 開始新 Sprint

### 3. 團隊協作

**交接工作**：
1. `sess-wrap` 存檔當前狀態
2. 將 `worklog/LastCheckPoint.log` 分享給隊友
3. 隊友使用 `sess-on` 接手

**定期同步**：
- 每日將 `worklog/` 和 `TODO.md` 推送到 git
- 團隊成員可查看最新進度

---

## 進階技巧

### 多專案管理

如果你同時進行多個專案：

```
project-a/
├── worklog/LastCheckPoint.log
└── TODO.md

project-b/
├── worklog/LastCheckPoint.log
└── TODO.md
```

切換專案時，Claude 會自動讀取對應的工作日誌。

### 自訂 Session 指令

你可以建立自己的 Session 指令：

**範例：sess-review（週回顧）**

`.claude/commands/sess-review.md`:
```markdown
請執行週回顧：

1. 讀取本週所有的 worklog/LastCheckPoint.log 記錄
2. 統計本週完成的項目
3. 分析進度是否符合預期
4. 提供下週建議

產生週報告到 shopfloor/Claude_TMP/etc/weekly-report-[日期].md
```

### Session 狀態視覺化

可以建立簡單的腳本查看狀態：

```bash
#!/bin/bash
# session-status.sh

echo "📊 Session 狀態"
echo "==============="
echo ""
echo "最後更新："
grep "最後更新" worklog/LastCheckPoint.log | head -1
echo ""
echo "完成度："
grep "完成度" worklog/LastCheckPoint.log | head -1
echo ""
echo "待辦事項："
grep -c "\[ \]" TODO.md
echo "項待完成"
```

---

## 疑難排解

### 問題 1：Session 指令沒有反應

**檢查**：
1. Slash Commands 是否正確設定（`.claude/commands/` 目錄）
2. 檔案是否存在（`worklog/LastCheckPoint.log`, `TODO.md`）

**解決**：
- 使用純文字指令：`Claude, sess on.`
- 或直接請 Claude：`請執行 Session 初始化`

### 問題 2：工作日誌格式錯誤

**檢查**：
- 確認 `worklog/LastCheckPoint.log` 格式正確
- 使用範例檔案重新初始化

**解決**：
```bash
cp examples/worklog/LastCheckPoint.log worklog/
```

### 問題 3：TODO.md 太雜亂

**建議**：
- 定期將已完成項目移到 `DONE.md`
- 使用優先級分類
- 每週清理過期項目

---

## 參考資料

- **SESSION_INIT.md** - Session 初始化清單
- **workflow-standards.md** - 工作流程規範
- **communication-standards.md** - 溝通規範

---

**最後更新**：2025-11-14
**版本**：2.0（通用版本）
