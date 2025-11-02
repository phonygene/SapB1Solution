# 專案待辦清單 (Project TODO List)

本檔案記錄所有「保留功能」、「未完成功能」、「偽代碼」的位置和相關資訊。

---

## 🔴 保留功能 (Reserved Features)

### 1. MDR 資料同步與程式呼叫
**狀態**: 保留（本地環境尚未建置）
**優先級**: 高
**說明**: 當費用申請單審核通過並成功寫入 SAP B1 AP Invoice 後，需要將營業稅發票資料同步到 MDR 資料庫，並呼叫 MDR 程式將資料寫入 SAP B1 營業稅外掛表。

**相關資訊**:
- MDR 資料庫名稱: `MDR`
- MDR 資料庫伺服器: `192.168.1.219`
- 目標表:
  - `MGUIAP_Import` (表頭)
  - `MGUIAPDetail_Import` (明細)
- 呼叫方式: 執行 .exe 程式
- 程式路徑: `//192.168.1.219/tools_shr/MDRImport.exe` (待確認)

**程式碼位置**:
- 檔案: `MgmSP/ExpenseClaimForm.aspx.vb` (或新的 AP Invoice 表單)
- 函式:
  - `WriteMDRData()` - 寫入 MDR 資料庫 (已預留偽代碼)
  - `CallMDRProgram()` - 呼叫 MDR .exe 程式 (已預留偽代碼，已註解)
- 行號: 待實作

**待確認**:
- [ ] MDR .exe 程式的完整路徑
- [ ] MDR .exe 程式的呼叫參數
- [ ] MDR .exe 程式的回傳值和錯誤處理方式
- [ ] 是否需要等待 .exe 執行完成？
- [ ] 是否需要記錄 .exe 執行日誌？

**預計實作時間**: TBD

---

## 🟡 未完成功能 (Incomplete Features)

### 1. jID 產生規則
**狀態**: 規格未定
**優先級**: 高
**說明**: 需要確定平台唯一單號 jID 的產生規則和格式。

**待確認**:
- [ ] jID 格式：純流水號 or 日期+流水號 or 其他？
- [ ] jID 產生方式：IDENTITY or Sequence or 自訂函式？
- [ ] jID 重置規則：是否每年/每月重置？

**程式碼位置**:
- 檔案: 待建立 `GetNextJID()` 函式
- 影響範圍: 所有需要產生 jID 的表單

---

## 🟢 已完成功能 (Completed Features)

### 1. 平台資料表建立
**完成日期**: 2025-10-30
**說明**: 建立 jOPCH, jPCH1, jMGUIAP, jMGUIAPDetail 四個表，包含 jID 欄位和所有必要索引。

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

## 🔍 快速查找

### 按檔案查找保留功能:
- `MgmSP/ExpenseClaimForm.aspx.vb`:
  - WriteMDRData() - MDR 資料同步
  - CallMDRProgram() - MDR 程式呼叫

### 按功能查找保留功能:
- MDR 相關: ExpenseClaimForm.aspx.vb
- jID 產生: 待建立

### 按優先級查找:
- 🔴 高優先級: MDR 資料同步、jID 產生規則
- 🟡 中優先級: （目前無）
- 🔵 低優先級（研究階段）: MCP 語音模組

---

**最後更新**: 2025-11-02
**維護者**: Claude + Jason
