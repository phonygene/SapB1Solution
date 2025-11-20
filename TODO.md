# 專案待辦清單 (Project TODO List)

本檔案記錄所有「保留功能」、「未完成功能」、「偽代碼」的位置和相關資訊。

---

## 📊 當前專案狀態（2025-11-18）

**專案結構**：雙線並行開發
- 🔴 **工作線 A**：費用申請單系統（ExpenseClaimForm）- 完成度 75%
- 🔵 **工作線 B**：SAP B1 升級對應表（MCP Server）- 完成度 60%
- 📊 **整體完成度**：70%

---

## 🔴 工作線 A：費用申請單系統

### 已完成項目 ✅

#### 1. jID 產生規則
**完成日期**: 2025-11-17
**狀態**: ✅ 已確定並實作
**決策**: 使用 IDENTITY 簡單流水號

**實作細節**:
- jOPCH.jID 使用 IDENTITY(1,1)
- 透過 SCOPE_IDENTITY() 取得新產生的 jID
- 檔案: `MgmSP/ExpenseClaimForm.aspx.vb`, SaveToJtdb_jOPCH() (Lines 395-457)

#### 2. jtdb 資料寫入邏輯
**完成日期**: 2025-11-17
**狀態**: ✅ 已實作（待架構調整）

**已實作方法**:
- SaveToJtdb_jOPCH() - 產生 jID (Lines 395-457)
- SaveToJtdb_jPCH1() - 寫入採購單明細 (Lines 464-516)
- SaveToJtdb_jMGUIAP() - 寫入費用申請單表頭 (Lines 523-566)
- SaveToJtdb_jMGUIAPDetail() - 寫入費用申請單明細 (Lines 573-643)
- UpdateJtdb_DocEntry() - 更新 DocEntry (Lines 650-701)

#### 3. 完整業務流程整合
**完成日期**: 2025-11-17
**狀態**: ✅ 已實作

**CreateExpenseClaim() 方法** (Lines 706-803):
- 第一步：寫入 jtdb 資料庫
- 第二步：建立 SAP AP Invoice
- 第三步：更新 DocEntry
- 第四步：同步至 MDR 資料庫
- 第五步：更新頁面顯示

#### 4. MDR 資料同步功能
**完成日期**: 2025-11-17
**狀態**: ✅ 已完整實作（原本為保留功能）

**WriteMDRData() 方法** (Lines 815-1080):
- 從 jtdb.jMGUIAP 和 jMGUIAPDetail 讀取資料
- 寫入 MDR.MGUIAP_Import 和 MGUIAPDetail_Import
- 完整的 Transaction 管理與錯誤處理
- MDR 同步失敗不影響主流程

**CallMDRProgram() 方法** (Lines 1089-1119):
- 狀態: 保留（待確認執行檔路徑與參數）

### 待辦事項（高優先級）⏳

#### 1. 建立費用項目對應表
**優先級**: 🔴 高（阻塞測試）
**說明**: 需要建立 expense_item_mapping 表，用於費用項目與會計科目的對應

**表結構**:
```sql
CREATE TABLE expense_item_mapping (
    ExpCode NVARCHAR(50) PRIMARY KEY,    -- 費用項目代碼
    ExpName NVARCHAR(100),                -- 費用項目名稱
    AcctCode NVARCHAR(50),                -- SAP 會計科目代碼
    CreateDate DATETIME DEFAULT GETDATE(),
    CreateBy NVARCHAR(50),
    UpdateDate DATETIME,
    UpdateBy NVARCHAR(50)
)
```

**待確認**:
- [ ] 初始資料有哪些常用費用項目？
- [ ] AcctCode 與 SAP B1 OACT 的對應關係？

#### 2. 架構調整：CardName 查詢
**優先級**: 🔴 高
**檔案**: MgmSP/ExpenseClaimForm.aspx.vb
**方法**: SaveToJtdb_jOPCH() (Line 438)

**問題**:
```vb.net
.AddWithValue("@CardName", cardCode)  ' 目前使用 CardCode
```

**解決方案**:
需要從 SAP B1 OCRD 查詢 CardName
```vb.net
' 查詢 CardName
Dim sqlCardName As String = "SELECT CardName FROM OCRD WHERE CardCode = @CardCode"
Dim cmdCardName As New SqlCommand(sqlCardName, SapConn)
cmdCardName.Parameters.AddWithValue("@CardCode", cardCode)
Dim cardName As String = CStr(cmdCardName.ExecuteScalar())
```

**待確認**:
- [ ] U_LIFNR 就是 SAP B1 的 CardCode？

#### 3. 架構調整：AcctCode 對應
**優先級**: 🔴 高
**檔案**: MgmSP/ExpenseClaimForm.aspx.vb
**方法**: SaveToJtdb_jPCH1() (Line 495), CreateExpenseClaim() (Line 749)

**問題**:
```vb.net
.AddWithValue("@ItemCode", "_SYS00000000001")  ' 錯誤：應使用 AcctCode
```

**解決方案**:
1. 從 expense_item_mapping 查詢 AcctCode
2. 使用 oInvoice.Lines.AccountCode（不是 ItemCode）

**待確認**:
- [ ] 費用申請單明細如何對應到費用項目？
- [ ] 是否需要在 GridView 增加 ExpCode 欄位？

#### 4. 架構調整：TaxCode 動態查詢
**優先級**: 🔴 高
**檔案**: MgmSP/ExpenseClaimForm.aspx.vb
**方法**: SaveToJtdb_jPCH1() (Line 500), CreateExpenseClaim() (Line 751)

**問題**:
```vb.net
.AddWithValue("@TaxCode", "100")  ' 固定值
oInvoice.Lines.TaxCode = "100"    ' 固定值
```

**解決方案**:
依 U_ZFORM_CODE 查詢 SAP B1 OVTG 取得對應的 TaxCode

**待確認**:
- [ ] U_ZFORM_CODE → OVTG 的對應邏輯？
- [ ] 是否需要建立 zform_tax_mapping 對應表？

#### 5. 使用者介面調整
**優先級**: 🟡 中
**檔案**: MgmSP/ExpenseClaimForm.aspx, ExpenseClaimForm.aspx.vb

**待調整**:
- [ ] GridView 增加費用項目下拉選單（綁定 expense_item_mapping）
- [ ] GridView 增加說明欄位文字輸入框
- [ ] 移除 ItemCode 相關欄位

#### 6. 完整測試流程
**優先級**: 🟡 中

**測試項目**:
- [ ] 編譯測試
- [ ] 單元測試（5 個 jtdb 寫入方法）
- [ ] 整合測試（完整流程：jtdb → SAP → MDR）
- [ ] 錯誤情境測試（SAP 建立失敗、MDR 同步失敗）

---

## 🔵 工作線 B：SAP B1 升級對應表

### 業務背景
- **時程**：2025 年 12 月切換到新版 SAP B1
- **開帳日**：10 月
- **需同步**：11 月交易與主檔（舊 → 新）
- **策略**：建立新舊對應表於 JTTST 資料庫

### 已完成項目 ✅

#### 1. MCP Server 多資料庫連線機制
**完成日期**: 2025-11-17
**狀態**: ✅ 已完成

**實作內容**:
- .env.jtdb - 生產環境設定（192.168.1.31 / jtdb）
- .env.JTTST - 測試環境設定（192.168.1.31 / JTTST / sa）
- 支援快速切換不同 SAP B1 資料庫

#### 2. 資料庫切換腳本
**完成日期**: 2025-11-17
**狀態**: ✅ 已完成

**檔案**:
- mcp-sqlserver/switch-db.sh（Linux）
- mcp-sqlserver/switch-db.bat（Windows）
- mcp-sqlserver/DATABASE_SWITCH_GUIDE.md（使用說明）

**使用方式**:
```bash
./switch-db.sh JTTST  # 切換到 JTTST
./switch-db.sh jtdb   # 切換回 jtdb
```

#### 3. 備份安全機制
**完成日期**: 2025-11-17
**狀態**: ✅ 已完成

**雙重保護**:
1. 環境變數層級：.env.JTTST 設定 BACKUP_ENABLED=false
2. 程式邏輯層級：config.json 設定 max_db_size_mb=100

**檔案**:
- mcp-sqlserver/config.json
- mcp-sqlserver/src/backup_manager.py

#### 4. 會話管理指令強化
**完成日期**: 2025-11-17
**狀態**: ✅ 已完成

**新增功能**:
- sess-off / sess-wrap 支援 -s / --selective 參數
- 選擇性記錄模式：使用者可勾選「不想記錄」的項目

**檔案**:
- .claude/commands/sess-off.md
- .claude/commands/sess-wrap.md
- collaboration-package/examples/.claude/commands/

### 待辦事項（高優先級）⏳

#### 1. 驗證 MCP Server 連線
**優先級**: 🔴 高
**說明**: 確認 MCP Server 正確連線至 JTTST 資料庫

**驗證步驟**:
- [ ] 重啟 Claude CLI 會話（確保讀取新 .env）
- [ ] 使用 mcp__sapb1-sql__get_db_status 檢查連線狀態
- [ ] 確認資料庫名稱為 JTTST
- [ ] 確認備份機制已停用

#### 2. 建立 9 個新舊對應表
**優先級**: 🔴 高
**說明**: 在 JTTST 資料庫建立新舊系統對應表

**對應表清單**:

| 資料表 | 用途 | 主鍵 |
|--------|------|------|
| mACT | 會計科目對應 | OC |
| mITM | 新舊料號對應 | OC |
| mITB | 項目群組對應 | OC |
| mCRD | 新舊業務夥伴代碼對應 | OC |
| mCRG | 業務夥伴群組對應 | OC |
| mVTG | 稅碼對應 | OC |
| mSLP | 銷售人員對應 | OC |
| mCRN | 幣別對應 | OC |
| mCTG | 付款條件對應 | OC |

**欄位說明**:
- OC (Old Code)：舊系統代碼（主鍵）
- ON (Old Name)：舊系統名稱
- NC (New Code)：新系統代碼
- NN (New Name)：新系統名稱

**標準結構**:
```sql
CREATE TABLE mXXX (
    OC NVARCHAR(50) NOT NULL PRIMARY KEY,
    ON NVARCHAR(200),
    NC NVARCHAR(50),
    NN NVARCHAR(200),
    CreateDate DATETIME DEFAULT GETDATE(),
    CreateBy NVARCHAR(50),
    UpdateDate DATETIME,
    UpdateBy NVARCHAR(50),
    Remark NVARCHAR(500)
)
```

**建立步驟**:
- [ ] 先建立 1-2 個表測試
- [ ] 驗證結構正確
- [ ] 批次建立其餘表

#### 3. 驗證對應表建立結果
**優先級**: 🔴 高

**驗證項目**:
- [ ] 使用 mcp__sapb1-sql__list_tables 確認表已建立
- [ ] 使用 mcp__sapb1-sql__get_table_info 檢查各表結構
- [ ] 確認主鍵設定正確
- [ ] 確認預設值設定正確

#### 4. 準備資料匯入腳本
**優先級**: 🟡 中

**待規劃**:
- [ ] 確認資料來源（舊系統 vs 新系統）
- [ ] 設計資料匯入流程
- [ ] 建立範例資料驗證
- [ ] 撰寫匯入 SQL 腳本或程式

#### 5. 建立對應表使用文件
**優先級**: 🟡 中

**待撰寫**:
- [ ] 記錄各對應表用途
- [ ] 提供查詢範例
- [ ] 說明資料維護流程
- [ ] 整合測試案例

---

## 🟢 已完成功能 (Completed Features)

### 1. 平台資料表建立
**完成日期**: 2025-10-30
**說明**: 建立 jOPCH, jPCH1, jMGUIAP, jMGUIAPDetail 四個表，包含 jID 欄位和所有必要索引。

### 2. jtdb 資料庫稽核欄位
**完成日期**: 2025-10-XX
**說明**: 在多個表增加稽核欄位（CreateDate, CreateBy, UpdateDate, UpdateBy）
- expense_category
- addr
- jPCH1
- jMGUIAPDetail
- jOPCH（審核欄位）

### 3. User 表 Approver 欄位
**完成日期**: 2025-10-XX
**說明**: 增加 Approver 欄位，8 位使用者設定為審核者

### 4. MDR 資料庫建立
**完成日期**: 2025-11-XX
**說明**: 建立 MGUIAP_Import 和 MGUIAPDetail_Import 表，包含索引和測試資料

---

## 📝 開發規範 (Development Guidelines)

### 保留功能的程式碼撰寫規範:
1. **盡量完整撰寫功能程式碼**，即使資訊不足也寫偽代碼
2. **使用註解標記保留功能**:
   ```vbnet
   ' ===== [保留功能] MDR 資料同步 =====
   ' 說明: 將營業稅發票資料同步到 MDR 資料庫
   ' 狀態: 本地環境尚未建置
   ' 預計實作日期: TBD
   ' TODO: 確認 MDR 資料庫連線字串和表結構
   Private Sub WriteMDRData(jID As Integer, docEntry As Integer)
       ' 偽代碼：連線到 MDR 資料庫
       ' Dim mdrConn As New SqlConnection("Server=192.168.1.219;Database=MDR;...")

       ' 偽代碼：寫入 MGUIAP_Import
       ' ...
   End Sub
   ```

3. **註解呼叫處**:
   ```vbnet
   ' 寫入 SAP B1 成功後，同步到 MDR
   ' [保留] 本地環境尚未建置
   ' WriteMDRData(jID, docEntry)
   ' CallMDRProgram()
   ```

4. **在本 TODO.md 中記錄**:
   - 功能說明
   - 程式碼位置（檔案、函式、行號）
   - 待確認的資訊
   - 預計實作時間

### 編碼風格標準:
- 優先使用 With 語句（節省 Token）
- 使用 GetCurrentUserId() 取得使用者
- 使用 $"..." 字串插值
- SQL Server 2005/2008 相容性（不使用 2008+ 新語法）

---

## 🔍 快速查找

### 按檔案查找:
**工作線 A - 費用申請單**:
- `MgmSP/ExpenseClaimForm.aspx.vb`:
  - SaveToJtdb_jOPCH() - Lines 395-457
  - SaveToJtdb_jPCH1() - Lines 464-516
  - SaveToJtdb_jMGUIAP() - Lines 523-566
  - SaveToJtdb_jMGUIAPDetail() - Lines 573-643
  - UpdateJtdb_DocEntry() - Lines 650-701
  - CreateExpenseClaim() - Lines 706-803
  - WriteMDRData() - Lines 815-1080
  - CallMDRProgram() - Lines 1089-1119（保留）

**工作線 B - MCP Server**:
- `mcp-sqlserver/config.json` - 備份大小限制設定
- `mcp-sqlserver/src/backup_manager.py` - 備份安全邏輯
- `mcp-sqlserver/switch-db.sh` - Linux 切換腳本
- `mcp-sqlserver/switch-db.bat` - Windows 切換腳本
- `.claude/commands/sess-off.md` - 會話結束指令
- `.claude/commands/sess-wrap.md` - 會話包裝指令

### 按優先級查找:
- 🔴 **高優先級**（工作線 A）:
  - 建立 expense_item_mapping 表
  - 架構調整：CardName 查詢
  - 架構調整：AcctCode 對應
  - 架構調整：TaxCode 動態查詢

- 🔴 **高優先級**（工作線 B）:
  - 驗證 MCP Server 連線至 JTTST
  - 建立 9 個對應表
  - 驗證對應表建立結果

- 🟡 **中優先級**:
  - 使用者介面調整
  - 完整測試流程
  - 準備資料匯入腳本
  - 建立對應表使用文件

---

## 🔵 研究階段功能 (Research & Future Features)

### 1. 語音優化協作 - MCP 語音模組
**狀態**: 草案討論中
**優先級**: 低（研究階段）
**預計工時**: 5-7 天（分階段）
**提出日期**: 2025-11-02

#### 功能概述
建立 MCP 語音模組，支援快捷鍵觸發語音輸入，自動轉換為文字後傳給 Claude Code。

#### 使用情境
1. **CTRL+F1**: 語音輸入 Slash Command
   - 範例: 說「sess off」→ 自動轉換為 `/sess-off`
   - 錄音模式: 單擊，3 秒或靜音自動結束

2. **CTRL+F2**: 語音輸入自然語言 Prompt
   - 範例: 說「請幫我列出最近三十次 commit 的紀錄...請開始執行」
   - 錄音模式: 雙擊啟動，檢測「請開始執行」或 2 秒靜音結束

3. **CTRL+F3**: 語音快速筆記
   - 直接追加到 TODO.md
   - 錄音模式: 單擊

#### 技術架構
**模組名稱**: `mcp-voice-assistant`

**技術棧**:
- Python（主程式）
- pynput（全域快捷鍵監聽）
- Azure Speech SDK / OpenAI Whisper（語音轉文字）
- MCP SDK（與 Claude Code 整合）

**組件**:
- HotkeyListener: 監聽全域快捷鍵
- VoiceRecorder: 錄製麥克風音訊
- SpeechToText: 串接 STT API
- MCPTools: MCP 工具介面

#### 技術待研究
- [ ] Claude Code 能否接受外部程式的輸入？
- [ ] MCP Server 能否主動推送訊息給 Claude？
- [ ] 語音轉文字 API 選型（Whisper vs Azure Speech）
- [ ] 全域快捷鍵監聽在 Windows 的實作方式
- [ ] 如何自動觸發 Slash Command 執行？

#### 階段性計畫
**Phase 1: 技術驗證（1-2 天）**
- [ ] 研究 Claude Code 的輸入機制
- [ ] 測試 MCP Server 能否主動推送訊息
- [ ] 驗證語音轉文字 API

**Phase 2: MVP - 剪貼簿方案（3-5 天）**
- [ ] 實作全域快捷鍵監聽
- [ ] 實作語音錄製與轉文字
- [ ] 實作基本功能：錄音 → 轉文字 → 複製到剪貼簿
- [ ] 使用者手動貼上（Ctrl+V）

**Phase 3: 自動化整合（視 Phase 1 結果）**
- [ ] 實作主動推送機制（如果可行）
- [ ] 或提供替代方案（剪貼簿、檔案）

**Phase 4: 優化體驗**
- [ ] 視覺/聽覺提示（錄音中、轉換中、完成）
- [ ] 錯誤處理（無法識別、網路錯誤）
- [ ] 配置介面（GUI 或 TUI）

#### 配置範例（config.yml）
```yaml
voice_assistant:
  stt_provider: "azure"  # azure / whisper / google
  stt_api_key: "${AZURE_SPEECH_KEY}"
  stt_region: "eastasia"

  hotkeys:
    - key: "ctrl+f1"
      mode: "slash_command"
      trigger: "single"

    - key: "ctrl+f2"
      mode: "natural_prompt"
      trigger: "double"
      end_phrase: "請開始執行"

    - key: "ctrl+f3"
      mode: "quick_note"
      trigger: "single"

  recording:
    max_duration: 60
    silence_timeout: 2
    language: "zh-TW"
```

#### 參考資料
- MCP SDK 文件
- pynput 快捷鍵範例
- Azure Speech Service
- OpenAI Whisper

#### 備註
2025-11-02 與 Claude 討論草案，建議先做剪貼簿 MVP 驗證可行性，待有空時再開始開發。

---

## 🔄 保留功能（已實作但需確認細節）

### 1. MDRImport.exe 程式呼叫
**狀態**: 程式碼已保留，待確認執行細節
**優先級**: 低
**檔案**: MgmSP/ExpenseClaimForm.aspx.vb (Lines 1089-1119)

**待確認**:
- [ ] MDR .exe 程式的完整路徑（目前推測：//192.168.1.219/tools_shr/MDRImport.exe）
- [ ] MDR .exe 程式的呼叫參數
- [ ] MDR .exe 程式的回傳值和錯誤處理方式
- [ ] 是否需要等待 .exe 執行完成？
- [ ] 是否需要記錄 .exe 執行日誌？

---

**最後更新**: 2025-11-18
**維護者**: Claude + Jason
