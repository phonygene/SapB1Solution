# 工具執行錯誤日誌

> **用途**：記錄 Agent 執行工具時遇到的錯誤，用於分析和避免重複錯誤
> **規則**：遇到工具執行錯誤（特別是重試 2 次以上仍失敗）時，必須記錄到此檔案

---

## 記錄格式

```markdown
### YYYY-MM-DD HH:MM | {工具名稱} | {錯誤類型}
- **命令/操作**: `實際執行的命令`
- **錯誤訊息**:
  ```
  錯誤訊息內容
  ```
- **根因分析**: 簡述為什麼會失敗
- **解決方案**: 如何解決（如果已解決）
- **狀態**: 已解決 / 待分析 / 已記錄到 skills
```

---

## 錯誤記錄

### 2026-01-08 14:30 | Claude Code Edit | Rust Panic（中文字符邊界）
- **命令/操作**: `Edit` 工具更新 `skills/INDEX.md`
- **錯誤訊息**:
  ```
  thread '<unnamed>' panicked at library\core\src\str\mod.rs:833:21:
  byte index 5 is not a char boundary; it is inside '範' (bytes 3..6) of `規範) |`
  ```
- **根因分析**: Claude Code 的 Rust 程式在處理中文字符時，使用 byte index 而非 char index 導致切到多位元組字符的中間
- **解決方案**: 這是 Claude Code 本身的 bug，無法由使用者解決。遇到時只能重試
- **狀態**: 待 Claude Code 官方修復（已回報 GitHub Issue）

### 2026-01-08 | Bash | MSBuild 路徑/參數錯誤
- **命令/操作**:
  ```bash
  "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe" /t:Build
  ```
- **錯誤訊息**:
  ```
  /usr/bin/bash: line 1: C:\Program Files...: No such file or directory
  Exit code 127
  ```
- **根因分析**:
  1. VS 版本錯誤：假設 `Community` 但實際是 `Professional`
  2. 路徑格式錯誤：Git Bash 無法識別 `C:\...` 格式
  3. 參數格式錯誤：Git Bash 把 `/t:` 的 `/` 解析為路徑
- **解決方案**:
  ```bash
  "/c/Program Files (x86)/Microsoft Visual Studio/2019/Professional/MSBuild/Current/Bin/MSBuild.exe" -t:Build
  ```
- **狀態**: 已記錄到 skills (`general-checklist.md`)

---

## 常見錯誤模式索引

| 錯誤模式 | 出現次數 | 解決方案位置 |
|----------|----------|--------------|
| Claude Code Rust Panic | 1 | 無解，重試或回報 GitHub |
| MSBuild 路徑/參數 | 1+ | `skills/general-checklist.md` |
| UTF-8 BOM 缺失 | 2+ | `skills/general-checklist.md` |
| COM 物件未釋放 | 多次 | `skills/backend-checklist.md` |

---

## 開發錯誤（非工具錯誤）

> 開發過程中的邏輯錯誤或疏忽，記錄以避免重複

### 2026-01-08 | PurchaseRequestForm | 搜尋彈窗事件混淆
- **問題描述**: 品號搜尋彈窗按搜尋按鈕後，結果變成供應商搜尋結果（使用表頭業務夥伴代碼搜尋）
- **影響範圍**: 請購單頁面的品號搜尋功能
- **可能根因**:
  1. 複製粘貼代碼時沒有仔細檢查事件綁定
  2. 兩個搜尋彈窗結構相似，容易混淆
  3. ModalPopupExtender 沒有設置 BehaviorID 可能導致行為衝突
  4. 開發時沒有參照費用申請單的實現模式
- **已採取措施**:
  1. 為 mpeVendor 和 mpeItem 添加明確的 BehaviorID
  2. 確認 ASPX 和 VB 代碼中的事件綁定正確
- **待確認**: 需要實際測試確認問題是否解決
- **教訓**:
  1. **開發時必須專注**：寫代碼時必須專注於當前功能，不能分心
  2. **參照範例**：有現成參考（費用申請單）時必須嚴格對照
  3. **測試完整流程**：新功能開發後必須測試完整流程，不能假設正確
  4. **BehaviorID 最佳實踐**：所有 ModalPopupExtender 都應設置 BehaviorID
- **狀態**: 待實測確認

---

## 待分析的錯誤

> 尚未找到解決方案的錯誤，等待分析

（目前無）
