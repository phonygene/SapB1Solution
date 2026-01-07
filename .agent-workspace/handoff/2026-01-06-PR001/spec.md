# 請購單功能規劃

## 任務 ID
2026-01-06-PR001

## 背景分析

### 現有架構
| 表格 | 用途 | 欄位數 |
|------|------|--------|
| jOPCH | 費用申請單表頭 | 42 |
| jPCH1 | 費用申請單明細 | 33 |
| jMGUIAP | 營業稅發票表頭 | - |
| jMGUIAPDetail | 營業稅發票明細 | - |

### 費用申請單 vs 請購單

| 項目 | 費用申請單 | 請購單 |
|------|-----------|--------|
| SAP 對應 | AP Invoice (OPCH) | Purchase Request (OPRQ) |
| 明細類型 | 服務型 (ExpCategory) | 項目型 (ItemCode) |
| 供應商 | 必填 | 可選（建議填寫） |
| 數量/單價 | 隱藏（用 LineTotal） | 顯示並計算 |
| 稅額計算 | 有 | 有 |
| 審核流程 | 有 | 有 |
| 後續單據 | 寫入 SAP AP Invoice | 轉採購訂單 (PO) |

### SAP 請購單特性
- SAP B1 目前無 OPRQ/PRQ1 表（未啟用模組）
- 平台需自建請購單表格
- 審核通過後可轉採購訂單

---

## 資料庫設計

### 表頭：jOPRQ（請購單表頭）

複用 jOPCH 結構，調整部分欄位：

```sql
CREATE TABLE [dbo].[jOPRQ] (
    -- 主鍵
    [jID] INT IDENTITY(1,1) PRIMARY KEY,

    -- SAP 對應（審核後轉 PO 才有值）
    [DocEntry] INT NULL,              -- 轉 PO 後的 DocEntry
    [DocNum] INT NULL,                -- 轉 PO 後的 DocNum

    -- 供應商（請購單為可選）
    [CardCode] NVARCHAR(50) NULL,     -- 建議供應商
    [CardName] NVARCHAR(100) NULL,

    -- 請購人/部門
    [ReqName] NVARCHAR(50) NOT NULL,  -- 請購人
    [ReqDept] NVARCHAR(50) NULL,      -- 請購部門

    -- 日期
    [DocDate] DATE NOT NULL,          -- 請購日期
    [ReqDate] DATE NULL,              -- 需求日期

    -- 幣別/匯率
    [DocCurrency] NVARCHAR(3) DEFAULT 'TWD',
    [DocRate] DECIMAL(19,6) DEFAULT 1,

    -- 金額
    [DocTotal] DECIMAL(19,6) DEFAULT 0,    -- 總計（含稅）
    [VatSum] DECIMAL(19,6) DEFAULT 0,      -- 稅額
    [DocTotalFC] DECIMAL(19,6) DEFAULT 0,  -- 外幣總計

    -- 備註
    [Comments] NVARCHAR(500) NULL,

    -- 狀態
    [DocStatus] NVARCHAR(1) DEFAULT 'O',   -- O=Open, C=Closed
    [Canceled] NVARCHAR(1) DEFAULT 'N',

    -- 審核
    [ApprovalStatus] NVARCHAR(20) DEFAULT 'Pending',
    [ApprovedBy] NVARCHAR(50) NULL,
    [ApprovedDate] DATETIME NULL,
    [ApprovalComments] NVARCHAR(500) NULL,

    -- 轉單狀態
    [ToPOStatus] NVARCHAR(1) DEFAULT 'N',  -- N=未轉, Y=已轉
    [ToPODate] DATETIME NULL,
    [ToPODocEntry] INT NULL,               -- 轉出的 PO DocEntry

    -- 外部簽核
    [U_PID] INT NULL,

    -- 稽核
    [CreateDate] DATETIME DEFAULT GETDATE(),
    [CreateBy] NVARCHAR(50) NULL,
    [UpdateDate] DATETIME NULL,
    [UpdateBy] NVARCHAR(50) NULL
);

-- 索引
CREATE INDEX IX_jOPRQ_CardCode ON jOPRQ(CardCode);
CREATE INDEX IX_jOPRQ_ReqName ON jOPRQ(ReqName);
CREATE INDEX IX_jOPRQ_DocDate ON jOPRQ(DocDate);
CREATE INDEX IX_jOPRQ_ApprovalStatus ON jOPRQ(ApprovalStatus);
CREATE INDEX IX_jOPRQ_ToPOStatus ON jOPRQ(ToPOStatus);
```

### 明細：jPRQ1（請購單明細）

複用 jPCH1 結構，強調項目型欄位：

```sql
CREATE TABLE [dbo].[jPRQ1] (
    -- 主鍵
    [jID] INT NOT NULL,
    [LineNum] INT NOT NULL,

    -- SAP 對應
    [DocEntry] INT NULL,
    [DocNum] INT NULL,

    -- 項目（必填）
    [ItemCode] NVARCHAR(50) NOT NULL,     -- 品號
    [Dscription] NVARCHAR(200) NULL,      -- 品名/說明

    -- 數量/單價（請購單核心）
    [Quantity] DECIMAL(19,6) NOT NULL DEFAULT 1,
    [OpenQty] DECIMAL(19,6) DEFAULT 0,    -- 未結數量
    [Price] DECIMAL(19,6) NOT NULL DEFAULT 0,
    [Currency] NVARCHAR(3) DEFAULT 'TWD',
    [Rate] DECIMAL(19,6) DEFAULT 1,

    -- 金額
    [LineTotal] DECIMAL(19,6) DEFAULT 0,  -- 未稅金額
    [GTotal] DECIMAL(19,6) DEFAULT 0,     -- 含稅金額

    -- 稅
    [VatGroup] NVARCHAR(20) NULL,
    [VatPrcnt] DECIMAL(19,6) DEFAULT 0,
    [LineVat] DECIMAL(19,6) DEFAULT 0,

    -- 倉庫/交期
    [WhsCode] NVARCHAR(50) NULL,          -- 倉庫
    [ShipDate] DATE NULL,                  -- 預計交期

    -- 成本中心
    [CostingCode] NVARCHAR(50) NULL,      -- 產品別
    [CostingCode2] NVARCHAR(50) NULL,     -- 部門
    [Project] NVARCHAR(50) NULL,          -- 專案

    -- 備註/附件
    [LineMemo] NVARCHAR(254) NULL,
    [Attachment] NVARCHAR(500) NULL,

    -- 狀態
    [LineStatus] NVARCHAR(1) DEFAULT 'O',

    -- 稽核
    [CreateDate] DATETIME DEFAULT GETDATE(),
    [CreateBy] NVARCHAR(20) NULL,
    [UpdateDate] DATETIME NULL,
    [UpdateBy] NVARCHAR(20) NULL,

    PRIMARY KEY (jID, LineNum),
    FOREIGN KEY (jID) REFERENCES jOPRQ(jID)
);

-- 索引
CREATE INDEX IX_jPRQ1_ItemCode ON jPRQ1(ItemCode);
CREATE INDEX IX_jPRQ1_WhsCode ON jPRQ1(WhsCode);
```

---

## 介面設計

### 佈局（與費用申請單一致）

```
┌─────────────────────────────────────────────────────┐
│                    請購單                            │
├─────────────────────────────────────────────────────┤
│                  工具列（按鈕）                       │
├──────────────────────┬──────────────────────────────┤
│   表頭左側區域        │   表頭右側區域                │
├──────────────────────┴──────────────────────────────┤
│              明細區域（GridView）                    │
├──────────────────────┬──────────────────────────────┤
│   表尾左側區域        │   表尾右側區域                │
└──────────────────────┴──────────────────────────────┘
```

### 表頭左側（6 欄位）

| # | 欄位 | 資料來源 | 必填 | 可編輯 |
|---|------|----------|------|--------|
| 1 | 請購人 | 當前使用者 | ✅ | ❌ |
| 2 | 請購部門 | User.Dept | ❌ | ✅ |
| 3 | 建議供應商代碼 | OCRD | ❌ | ✅ |
| 4 | 建議供應商名稱 | OCRD | ❌ | ✅ |
| 5 | 幣別 | OCRN | ✅ | ✅ |
| 6 | 匯率 | ORTT | ✅ | 條件 |

### 表頭右側（5 欄位）

| # | 欄位 | 說明 | 必填 | 可編輯 |
|---|------|------|------|--------|
| 1 | 請購單號 | jID | ❌ | ❌ |
| 2 | 請購日期 | 預設當日 | ✅ | ✅ |
| 3 | 需求日期 | 預計需要日期 | ❌ | ✅ |
| 4 | 審核狀態 | Pending/Approved/Rejected | ❌ | ❌ |
| 5 | 簽核 PID | 外部簽核系統 | ❌ | ✅ |

### 明細區域（10 欄位）

| # | 欄位 | 說明 | 必填 | 可編輯 |
|---|------|------|------|--------|
| 1 | # | 列號 | ✅ | ❌ |
| 2 | 品號 | ItemCode（查詢子視窗） | ✅ | ✅ |
| 3 | 品名 | 自動帶入，可編輯 | ✅ | ✅ |
| 4 | 數量 | Quantity | ✅ | ✅ |
| 5 | 單價 | Price | ✅ | ✅ |
| 6 | 稅碼 | VatGroup（下拉） | ✅ | ✅ |
| 7 | 稅額 | 自動計算 | ❌ | ❌ |
| 8 | 含稅金額 | 自動計算 | ❌ | ❌ |
| 9 | 倉庫 | WhsCode（下拉） | ❌ | ✅ |
| 10 | 交期 | ShipDate | ❌ | ✅ |

### 表尾左側（2 欄位）

| # | 欄位 | 說明 |
|---|------|------|
| 1 | 採購人員 | SlpCode（下拉） |
| 2 | 備註 | Comments |

### 表尾右側（金額區）

| # | 欄位 | 說明 |
|---|------|------|
| 1 | 未稅總計 | SUM(LineTotal) |
| 2 | 稅額總計 | SUM(LineVat) |
| 3 | 含稅總計 | DocTotal |

---

## 業務流程

```
1. 使用者建立請購單
   ↓
2. 填寫品項、數量、單價
   ↓
3. 儲存（狀態: Pending）
   ↓
4. 匯出 PDF → 外部簽核
   ↓
5. 簽核通過 → 回填 PID
   ↓
6. 財會審核
   ├─ 駁回 → 修改或刪除
   └─ 放行 → 可轉採購訂單
       ↓
7. 轉採購訂單（Phase 2）
   ↓
8. 寫入 SAP PO
```

---

## 實作階段

### Phase 1：基礎功能
- [ ] 建立 jOPRQ、jPRQ1 資料表
- [ ] 建立 PurchaseRequestForm.aspx 介面
- [ ] 實作 CRUD 功能
- [ ] 實作品號查詢子視窗
- [ ] 實作金額自動計算
- [ ] 實作審核流程

### Phase 2：進階功能
- [ ] 轉採購訂單功能
- [ ] SAP PO API 整合
- [ ] 請購單查詢/列表頁面
- [ ] PDF 匯出功能

### Phase 3：優化
- [ ] 複製單據功能
- [ ] 批次審核
- [ ] 報表功能

---

## 與費用申請單差異摘要

| 項目 | 費用申請單 | 請購單 |
|------|-----------|--------|
| 表頭表 | jOPCH | jOPRQ |
| 明細表 | jPCH1 | jPRQ1 |
| 明細核心 | ExpCategory | ItemCode |
| 數量欄位 | 隱藏 | 顯示 |
| 單價欄位 | 隱藏 | 顯示 |
| 供應商 | 必填 | 可選 |
| 後續單據 | AP Invoice | Purchase Order |
| 營業稅頁籤 | 有 | 無 |

---

## 待確認事項（已確認）

| # | 問題 | 決定 |
|---|------|------|
| 1 | 品號來源 | SAP OITM |
| 2 | 倉庫來源 | SAP OWHS |
| 3 | 轉 PO 時機 | 審核者手動放行（與費用申請單一致） |
| 4 | 多階審核 | 不需要 |
| 5 | 預算控管 | 不需要 |

**確認日期**：2026-01-07

---

**建立時間**：2026-01-06
**規劃者**：Manager Agent
