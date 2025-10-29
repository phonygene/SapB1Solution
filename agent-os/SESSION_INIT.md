# Claude 協作初始化清單

**版本**：1.0
**更新日期**：2025-10-29

---

## 每次新對話開始時的初始化步驟

當使用者說 `Claude, sess on.` 時，Claude 必須依序執行：

### 第一步：讀取核心規範
按優先順序讀取以下檔案：

1. `agent-os/standards/global/session-management.md` - 了解 session 管理機制
2. `agent-os/standards/global/communication-standards.md` - 溝通與責任歸屬規範
3. `agent-os/standards/global/localization.md` - 語言與術語標準
4. `agent-os/standards/global/coding-style.md` - 程式碼風格規範
5. `shopfloor/Claude_TMP/etc/README_協作模式說明.txt` - 檔案輸出協作方式

### 第二步：讀取專案狀態
1. `worklog/LastCheckPoint.log` - 最新工作狀態與待辦事項
2. `對話狀態_*.md`（如果有最新的） - 階段性狀態摘要

### 第三步：報告與確認
向使用者報告：
- 上次工作的時間點
- 當前專案狀態摘要（完成度百分比）
- 未完成的待辦事項（按優先順序）
- **立即需要確認的事項**（如：等待使用者執行的任務）
- 建議的下一步工作

### 第四步：等待使用者回應
- 使用者可能回報上次待辦事項的執行結果
- 使用者可能提出新的任務
- 根據回應調整工作計畫

---

## 專案基本資訊（快速參考）

### 專案名稱
SapB1Solution - 費用申請單功能開發

### 當前階段
5 天衝刺計畫（2025-10-29 至 2025-11-02）
- 目標：完成費用申請單 Create 模式
- 架構：B1Transaction v1.1 規格
- 功能：AP Invoice + MDR 發票明細

### 技術環境
- **語言**：VB.NET WebForms (.NET Framework 4.0)
- **IDE**：Visual Studio（使用者端）
- **資料庫**：
  - 本機：jtdb (.\SQLEXPRESS2008R2)
  - SAP：JTTST1 (192.168.1.31)
- **API**：SAP Business One DI API
- **套件**：Newtonsoft.Json、iTextSharp、ExcelDataReader

### 專案結構
```
MgmSP/                    # 主要專案目錄
├── commcode/             # 共用程式碼
│   ├── CommUtil.vb       # 工具類別（SQL、DI API）
│   └── CommSignOff.vb    # 簽核相關
├── usermgm/              # 使用者管理
├── signoff/              # 簽核功能
└── [其他模組]/

shopfloor/Claude_TMP/     # Claude 產出檔案
├── SqlQuery/             # SQL 腳本
├── dNet/                 # VB.NET 檔案
└── etc/                  # 文件與配置

agent-os/                 # 協作規範
├── standards/global/     # 全域標準
└── config.yml            # Agent-OS 配置
```

### 溝通語言與術語
- **語言**：繁體中文（Traditional Chinese）
- **Row** = 列（lié）
- **Column** = 欄（lán）
- **時區**：台灣 UTC+8

### 協作模式
- **檔案輸出**：產生檔案到 `shopfloor/Claude_TMP/`，不在對話中貼大段程式碼
- **溝通方式**：簡要說明 + 等待使用者回報結果
- **回應結尾**：加上台灣時間標籤 `[YYYY-MM-DD HH:mm]`

---

## 重要規範摘要

### Session 管理指令
- `Claude, sess on.` - 上班（讀取進度）
- `Claude, sess check.` - 查看進度（不寫檔案）
- `Claude, sess wrap.` - 階段存檔，繼續工作
- `Claude, sess off.` - 完整存檔並下班

### 責任歸屬
- ✅ 正確：「抱歉，我之前提供的程式碼有誤...」
- ❌ 錯誤：「看來是你在某某地方弄錯了...」

### 方案呈現
- 主要方案：詳細說明（程式碼 + 步驟）
- 其他方案：簡要描述 + 優缺點對比

---

## 檔案命名規範（Claude_TMP 產出）

### SQL 腳本
格式：`[編號]_[用途]_[物件名稱].sql`
範例：`01_CreateTable_user_vendor_taxid.sql`

### VB.NET 檔案
格式：`[類別名稱].vb` 或 `[頁面名稱].aspx`
範例：`UserVendorTaxId.vb`、`ExpenseClaimForm.aspx`

### JSON 配置
格式：`[物件名稱]Config.json`
範例：`ExpenseClaimTransactionConfig.json`

---

## 快速檢查：Claude 是否正確初始化？

如果 Claude 正確讀取並理解了所有規範，應該能回答：

- [ ] 當前專案的目標是什麼？（費用申請單功能）
- [ ] 使用什麼語言溝通？（繁體中文）
- [ ] Row 和 Column 的正確譯法？（列、欄）
- [ ] 檔案應該產生到哪裡？（shopfloor/Claude_TMP/）
- [ ] 當前是 5 天衝刺的第幾天？（查看 LastCheckPoint.log）
- [ ] 上次工作未完成的待辦事項？（查看 LastCheckPoint.log）

---

**使用方式**：

每次新對話時，輸入：
```
Claude, sess on. 讀取 agent-os/SESSION_INIT.md 並執行初始化。
```

或更簡潔（假設 Claude 已熟悉流程）：
```
Claude, sess on.
```

Claude 會自動讀取本檔案並執行所有初始化步驟。

---

**維護建議**：

- 當協作規範有重大變更時，更新本檔案的「專案基本資訊」區塊
- 當專案進入新階段時，更新「當前階段」資訊
- 保持本檔案簡潔，詳細規範仍在各自的 .md 檔案中
