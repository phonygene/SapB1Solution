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

---

## 檔案編碼強制規則

> 🔴 **這是重複發生的問題！2026-01-07 和 2026-01-14 都因此出錯。**
>
> **Claude Code 的 Write/Edit 工具不會保留 BOM，每次寫入後必須驗證！**

### 情境一：創建新 .aspx/.vb 檔案

**必須立即執行以下步驟，不可跳過。**

```powershell
# 步驟 1：轉換為 UTF-8 with BOM
$path = "路徑"
$content = Get-Content -Path $path -Raw -Encoding UTF8
$utf8Bom = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllText($path, $content, $utf8Bom)

# 步驟 2：驗證 BOM（應輸出 239,187,191）
[System.IO.File]::ReadAllBytes($path)[0..2] -join ','
```

### 情境二：修改現有 .aspx/.vb 檔案（2026-01-14 新增）

**修改後必須驗證編碼未被改變！**

```powershell
# 驗證 BOM 仍存在（應輸出 239,187,191）
[System.IO.File]::ReadAllBytes("檔案路徑")[0..2] -join ','

# 如果不是 239,187,191，執行修復
$path = "檔案路徑"
$content = Get-Content -Path $path -Raw -Encoding UTF8
$utf8Bom = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllText($path, $content, $utf8Bom)
```

### 情境三：Git commit 前

**建議執行編碼檢查（或依賴 pre-commit hook）**

```bash
# 檢查即將 commit 的 .aspx/.vb 檔案編碼
git diff --cached --name-only | grep -E '\.(aspx|vb)$'
```

### 檢查清單

- [ ] 確認 .aspx.designer.vb 存在且已更新
- [ ] 確認 Inherits 有 `MgmSP.` 前綴
- [ ] 確認編碼為 UTF-8 with BOM（239,187,191）

### 為什麼這很重要？

- 缺少 BOM → 中文變亂碼 → 伺服器剖析錯誤 → 頁面無法執行
- 編碼被改變 → 所有 Unicode 符號變成 ?? → 介面損壞
- 這個問題已經發生多次，必須從流程上杜絕

## 程式碼品質

- [ ] 有適當的錯誤處理（Try-Catch）
- [ ] 敏感操作有記錄到 EventLog
- [ ] 不使用 magic number，用常數定義

---

## 代碼可靠性（Robustness）

> 防禦性程式設計原則，確保代碼在非預期情況下仍能安全運作

### 輸入驗證

- [ ] 所有外部輸入（用戶、API、檔案）都要驗證
- [ ] 數值範圍檢查（最小值、最大值、溢位）
- [ ] 字串長度與格式驗證
- [ ] Null/Empty 檢查在使用前執行

### 邊界條件

- [ ] 空集合/空陣列處理
- [ ] 零值、負值、極大值測試
- [ ] 首項/末項特殊處理
- [ ] 並發/競爭條件考量

### 錯誤處理

- [ ] 具體的例外類型（避免 catch Exception）
- [ ] 有意義的錯誤訊息（含上下文）
- [ ] 錯誤後的狀態清理（Finally/Using）
- [ ] 不吞掉例外（至少記錄）

### 資源管理

- [ ] 使用 Using/Try-Finally 確保釋放
- [ ] 資料庫連線、檔案句柄、COM 物件
- [ ] 避免資源洩漏（記憶體、連線池）

### 防禦性程式設計

- [ ] 假設外部資料不可信
- [ ] 假設網路/服務可能失敗
- [ ] 關鍵操作前檢查前置條件
- [ ] 避免 magic number，使用常數

---

## 工具錯誤處理規範（絕對遵守）

> **重複犯同樣的錯誤是最大的浪費。遇到錯誤必須記錄，避免自己和其他 Agent 重蹈覆轍。**

### 錯誤記錄觸發條件

以下情況**必須**記錄到 `work-logs/insights/tool-errors.md`：

1. **重試 2 次以上仍失敗**的工具執行
2. **Exit code 非 0** 且錯誤訊息不明顯
3. **路徑/環境相關**的錯誤（這類錯誤會重複發生）
4. **第一次遇到**的新類型錯誤

### 記錄內容

```markdown
### YYYY-MM-DD HH:MM | {工具名稱} | {錯誤類型}
- **命令/操作**: `實際執行的命令`
- **錯誤訊息**: (貼上錯誤內容)
- **根因分析**: 為什麼失敗
- **解決方案**: 如何解決
- **狀態**: 已解決 / 待分析 / 已記錄到 skills
```

### 解決錯誤後

1. **立即更新** `tool-errors.md` 的解決方案
2. **如果是通用問題**：同步更新到對應的 `skills/*.md`
3. **更新索引**：在 `tool-errors.md` 的「常見錯誤模式索引」新增條目

### 執行任務前（建議）

遇到以下類型的操作時，**先查閱** `work-logs/insights/tool-errors.md`：
- MSBuild / 編譯
- 檔案編碼操作
- Windows 路徑操作
- 外部工具呼叫

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

## 防呆/防錯機制的系統性實作原則（絕對遵守）

> **紅綠燈原則**：發現一個路口需要紅綠燈時，必須在所有路口都裝設。

### 核心原則

當發現某個防呆/防錯機制的需求時：

1. **全面盤點**：列出所有可能需要相同機制的位置
2. **與使用者確認**：呈現完整清單，確認實作範圍
3. **統一實作**：一次性完成所有位置的修改
4. **架構優化**：如果數量過多，考慮從架構層面解決（封裝成共用函數/模組）

### 授權原則

#### 標準修改（同類紅綠燈）
- 使用者同意方案後 → 一次性授權，不需逐個確認
- AI 自行判斷各位置的具體實作
- 完成後彙報總結

#### 新設計/變種（不同類型號誌）
- **需要討論**：設計原則、適用情境、各變種的使用條件
- **原則確定後** → AI 自行判斷具體實作
- 有困惑時才回報討論

#### 免確認情況

以下操作不需請求使用者同意：

1. **Agent 自身狀態變更**
   - 更新自己的狀態檔案（如 `current.md` 的狀態欄位）
   - 未來任何其它狀態機制的自我更新

2. **`*jdi*` 指令**（just do it）
   - 指令開頭為 `*jdi*` 時，執行該動作不需請求同意
   - 用途：Manager 指派已確認過的任務腳本時，避免溝通被同意請求打斷
   - 前提：任務內容腳本已與使用者確認過

### 不可逆操作的保護

#### Git 備份
- 嚴守 commit 規範，確保所有修改可還原

#### 刪除/不可逆/無紀錄操作
- **數量小**：每次確認
- **數量大**：列出批次清單，一次確認後執行

### 禁止的做法

❌ 發現問題 → 只修一處 → 下次出問題再修另一處（逐個修補）

### 正確的做法

✅ 發現問題 → 盤點所有同類位置 → 確認範圍 → 統一實作或架構優化

### 架構優化的時機

當同類位置超過 3 處時，應考慮：
- 封裝成共用函數（如 `safe_send_key()`）
- 使用裝飾器/中介層統一處理
- 建立專用的配置/模板機制

### 新功能開發時的要求

如同道路規劃必須同時規劃：人行道、車道、號誌、管線、預留空間

新功能開發時必須同時考慮：
- 錯誤處理機制
- 日誌記錄
- 狀態驗證
- 回退/復原方案
- 與現有系統的整合點

---

## Windows 環境下的 Bash 執行規範

> Claude Code 的 Bash 工具在 Windows 上使用 Git Bash (Unix shell)，需遵循以下規則

### MSBuild 執行（編譯專案）

**正確寫法**：
```bash
"/c/Program Files (x86)/Microsoft Visual Studio/2019/Professional/MSBuild/Current/Bin/MSBuild.exe" \
  "C:/Projects/SapB1Solution/MgmSP/MgmSP.vbproj" \
  -t:Build \
  -p:Configuration=Debug \
  -verbosity:minimal
```

**常見錯誤**：

| 錯誤 | 說明 |
|------|------|
| `2019/Community/...` | ❌ 本機是 `2019/Professional/...` |
| `/t:Build` | ❌ Git Bash 會把 `/` 解析為路徑，應使用 `-t:Build` |
| `C:\Program Files\...` | ❌ 需用 `/c/Program Files/...` 或引號包裹 |

**可用的 MSBuild 路徑（本機）**：
- VS 2019: `/c/Program Files (x86)/Microsoft Visual Studio/2019/Professional/MSBuild/Current/Bin/MSBuild.exe`
- VS 2022: `/c/Program Files (x86)/Microsoft Visual Studio/2022/BuildTools/MSBuild/Current/Bin/MSBuild.exe`

### 執行 Windows 程式的通用規則

1. **路徑格式**：使用 `/c/...` 或用雙引號包裹 `"C:/..."`
2. **參數前綴**：使用 `-` 而非 `/`（避免被解析為路徑）
3. **空格路徑**：整個路徑用雙引號包裹

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

### 2026-01-06: 修飾鍵釋放機制不完整
- **問題**: pyautogui 按鍵操作後修飾鍵可能殘留，導致 Enter 變成 Shift+Enter（換行）
- **根因**: 釋放修飾鍵的代碼分散在多處，且有遺漏（typewrite 後無釋放）
- **解法**: 建立 `_release_modifiers()` 統一函數，在所有按鍵操作後調用
- **教訓**: 應用「紅綠燈原則」—發現防呆需求時，盤點所有同類位置統一實作

### 2026-01-08: MSBuild 執行失敗
- **問題**: Agent 使用錯誤的 VS 路徑 (`Community` 而非 `Professional`) 和參數格式 (`/t:` 而非 `-t:`)
- **根因**:
  1. Agent 假設 VS 安裝版本，未先驗證實際路徑
  2. 不了解 Git Bash 會將 `/` 解析為路徑
- **解法**: 在 general-checklist.md 新增「Windows 環境下的 Bash 執行規範」區塊
- **教訓**:
  1. Windows 路徑和參數在 Git Bash 中需要特殊處理
  2. 外部工具路徑應先驗證，不要假設

### 2026-01-07: UTF-8 BOM 問題再次發生
- **問題**: 新建的 PurchaseRequestForm.aspx 缺少 BOM，導致中文變亂碼、頁面剖析失敗
- **根因**:
  1. Claude Code 的 Write 工具預設不添加 BOM
  2. 規範存在但只是「被動檢查清單」，不是「主動操作步驟」
  3. 沒有「創建新檔案後必做」的強制流程
- **解法**: 在 general-checklist.md 新增「創建新 .aspx/.vb 檔案後（強制執行）」區塊
- **教訓**:
  1. 知道規範 ≠ 執行規範，必須有明確的操作步驟
  2. 工具的限制必須用流程來彌補
  3. 重複犯錯表示規範不夠具體，需要升級為強制流程

### 2026-01-14: UTF-8 BOM 問題第三次發生（修改現有檔案）
- **問題**: PurchaseRequestForm.aspx 編碼從 UTF-8 變成 Big5，所有 Unicode 符號變成 ??
- **根因**:
  1. 規範只覆蓋「創建新檔案」，沒有覆蓋「修改現有檔案」
  2. Claude Code 的 Write/Edit 工具可能改變現有檔案的編碼
  3. 沒有 commit 前的自動檢查機制
- **解法**:
  1. 更新 CLAUDE.md 加入顯眼的編碼規則
  2. 擴展 general-checklist.md 覆蓋「修改現有檔案」情境
  3. 新增 pre-commit hook 自動檢查編碼
- **教訓**:
  1. 規範必須覆蓋所有情境，不只是最常見的
  2. 被動規範不可靠，需要自動化防呆機制
  3. 同類問題第三次發生 → 必須加入自動化阻擋
