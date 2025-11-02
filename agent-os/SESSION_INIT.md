# Claude 協作初始化清單

**版本**：1.2
**更新日期**：2025-11-02

---

## 說明

本檔案是 Session 初始化的**執行清單**。
詳細的 Session 管理規範（指令定義、檔案格式）請參考 `agent-os/standards/global/session-management.md`。

---

## Session On 初始化步驟

當觸發 `/sess-on` 指令時，依序執行：

### 第一步：讀取核心規範
按優先順序讀取以下檔案：

1. `agent-os/standards/global/session-management.md` - Session 管理機制
2. `agent-os/standards/global/workflow-standards.md` - 工作流程與 Shopfloor 協作模式
3. `agent-os/standards/global/communication-standards.md` - 溝通與責任歸屬規範
4. `agent-os/standards/global/localization.md` - 語言與術語標準
5. `agent-os/standards/global/coding-style.md` - 程式碼風格規範
6. `shopfloor/Claude_TMP/etc/README_協作模式說明.txt` - 檔案輸出協作方式（補充說明）

### 第二步：讀取專案狀態
1. `worklog/LastCheckPoint.log` - 最新工作狀態與待辦事項
2. `TODO.md` - 保留功能與未完成項目

### 第三步：報告與確認
向使用者報告：
- 上次工作的時間點
- 當前專案狀態摘要（完成度百分比）
- 未完成的待辦事項（按優先順序）
- **立即需要確認的事項**（如：等待使用者執行的任務）
- 建議的下一步工作
- **⚠️ 協作模式提醒**：提醒使用者本專案採用 Shopfloor 協作模式

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

## 使用方式

### Slash Commands（推薦）
使用 Claude Code 的 slash command 功能：
- `/sess-on` - 上班/開始工作
- `/sess-check` - 查看進度（不寫檔案）
- `/sess-wrap` - 階段存檔，繼續工作
- `/sess-off` - 完整存檔並下班

### 純文字指令（備用）
如果 slash command 無法使用，可以輸入：
- `Claude, sess on.`
- `Claude, sess check.`
- `Claude, sess wrap.`
- `Claude, sess off.`

---

## 維護建議

- 當協作規範有重大變更時，更新本檔案的「專案基本資訊」區塊
- 當專案進入新階段時，更新「當前階段」資訊
- 保持本檔案簡潔，詳細規範仍在各自的 .md 檔案中
- Session 指令的詳細定義請維護在 `agent-os/standards/global/session-management.md`
