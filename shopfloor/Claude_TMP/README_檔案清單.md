# 費用申請單 - 完整檔案清單

**建立日期**: 2025-11-05
**狀態**: ✅ 所有檔案已產生完成

---

## 📂 檔案結構

```
shopfloor/Claude_TMP/
├── SqlQuery/                                    # SQL 腳本資料夾
│   ├── 05_CreateTable_addr.sql                 # 建立收貨地址表
│   ├── 06_CreateTable_expense_category.sql     # 建立費用類別表
│   ├── 07_AlterTable_jOPCH_Add_ApprovalComments.sql  # 新增審核意見欄位
│   └── 08_AlterTable_User_Add_CanApproveExpense.sql  # 新增審核權限欄位
│
├── ExpenseClaimForm_Part1_Header.aspx          # 第一階段：表頭架構（完整檔案）
├── ExpenseClaimForm_Part1_Header.aspx.vb       # 第一階段：表頭 CodeBehind（完整檔案）
│
├── ExpenseClaimForm_Part2_GridView.aspx.snippet       # 第二階段：費用明細 GridView（片段）
├── ExpenseClaimForm_Part2_GridView.aspx.vb.snippet    # 第二階段：費用明細 CodeBehind（片段）
│
├── ExpenseClaimForm_Part3_MDR_Tab.aspx.snippet        # 第三階段：MDR Tab（片段）
├── ExpenseClaimForm_Part3_MDR_Tab.aspx.vb.snippet     # 第三階段：MDR Tab CodeBehind（片段）
│
├── 整合說明_ExpenseClaimForm.md                # 完整整合說明文件
├── Web.config.ExpenseClaim.example             # Web.config 配置範例
└── README_檔案清單.md                          # 本檔案
```

---

## 📋 檔案說明

### 1. SQL 腳本（4 個檔案）

所有 SQL 腳本位於 `SqlQuery/` 資料夾：

| 檔名 | 用途 | 執行順序 |
|------|------|----------|
| 05_CreateTable_addr.sql | 建立收貨地址主檔表 | 1 |
| 06_CreateTable_expense_category.sql | 建立費用類別主檔表（含測試資料） | 2 |
| 07_AlterTable_jOPCH_Add_ApprovalComments.sql | jOPCH 表新增審核意見欄位 | 3 |
| 08_AlterTable_User_Add_CanApproveExpense.sql | User 表新增審核權限欄位 | 4 |

**執行方式**：
- 使用 SSMS 連線到 `jtdb` 資料庫
- 依序執行上述 4 個 SQL 檔案
- 或使用 Claude 的 MCP SQL 工具執行

---

### 2. 介面檔案（6 個檔案）

#### 第一階段：表頭與基本架構

| 檔名 | 類型 | 說明 |
|------|------|------|
| ExpenseClaimForm_Part1_Header.aspx | 完整檔案 | 包含表頭欄位、審核區塊、Tab 架構、CSS 樣式 |
| ExpenseClaimForm_Part1_Header.aspx.vb | 完整檔案 | 包含表頭邏輯、檔案上傳、審核功能、郵件通知 |

**功能**：
- ✅ 供應商選擇（連動聯絡人）
- ✅ 收貨地址、請款日期、到期日
- ✅ 產品/部門、幣別、匯率（自動更新）
- ✅ 備註、附件上傳/下載
- ✅ 審核區塊（放行/駁回/發送通知）
- ✅ Tab 切換（費用申請 / MDR）
- ✅ 儲存/送出/刪除/取消按鈕

#### 第二階段：費用明細 GridView

| 檔名 | 類型 | 說明 |
|------|------|------|
| ExpenseClaimForm_Part2_GridView.aspx.snippet | ASPX 片段 | GridView 介面（工具列、表格、合計區） |
| ExpenseClaimForm_Part2_GridView.aspx.vb.snippet | VB 片段 | GridView 程式邏輯（新增/刪除/計算） |

**功能**：
- ✅ 新增/刪除明細行
- ✅ 費用類別下拉選單（自動填入總帳科目）
- ✅ 費用說明、數量、單價
- ✅ 自動計算外幣金額、本幣金額
- ✅ 合計顯示（外幣總額、本幣總額）
- ✅ 資料驗證（必填欄位、數量 > 0）

#### 第三階段：MDR 發票明細 Tab

| 檔名 | 類型 | 說明 |
|------|------|------|
| ExpenseClaimForm_Part3_MDR_Tab.aspx.snippet | ASPX 片段 | MDR Tab 介面（表頭同步、發票明細） |
| ExpenseClaimForm_Part3_MDR_Tab.aspx.vb.snippet | VB 片段 | MDR Tab 程式邏輯（同步/驗證） |

**功能**：
- ✅ 表頭資訊即時同步（唯讀）
- ✅ 發票明細 GridView（新增/刪除）
- ✅ 發票號碼、日期、金額（未稅）、稅額
- ✅ 自動計算發票總額（含稅）
- ✅ 驗證金額總和（與 AP 發票總額比對）
- ✅ 驗證結果顯示（成功/失敗）

---

### 3. 說明文件（2 個檔案）

| 檔名 | 用途 |
|------|------|
| 整合說明_ExpenseClaimForm.md | **最重要**：完整的整合步驟、測試流程、常見問題排查 |
| Web.config.ExpenseClaim.example | Web.config 配置範例（連線字串、SMTP、檔案上傳） |

---

## 🚀 快速開始（3 步驟）

### 步驟 1：執行 SQL 腳本

```bash
# 依序執行 SqlQuery/ 資料夾中的 4 個 SQL 檔案
# 執行完成後，設定審核權限：
# UPDATE [User] SET CanApproveExpense = 1 WHERE id = 'admin'
```

### 步驟 2：整合程式碼

請詳細閱讀 `整合說明_ExpenseClaimForm.md`，依照以下順序整合：

1. 複製 Part1 檔案到專案目錄
2. 插入 Part2 片段（費用明細 GridView）
3. 插入 Part3 片段（MDR Tab）
4. 修改現有函式（Page_Load、btnSave_Click 等）
5. 配置 Web.config

### 步驟 3：測試

```bash
# 建立上傳資料夾
mkdir -p shopfloor/Uploads/ExpenseClaims

# 編譯專案
# 測試新增單據、檔案上傳、明細計算、MDR 驗證、審核功能
```

---

## ⚙️ 技術規格

### 技術棧
- **前端**: ASP.NET WebForms 4.0
- **後端**: VB.NET
- **資料庫**: SQL Server（jtdb + JTTST1）
- **郵件**: SMTP (smg.jettech.com.tw)

### 資料表
- `jOPCH` - 費用申請單表頭
- `jOPC1` - 費用申請單明細
- `jMDR1` - MDR 發票明細
- `addr` - 收貨地址主檔
- `expense_category` - 費用類別主檔
- `User` - 使用者表（新增 CanApproveExpense 欄位）

### 外部相依
- SAP B1 資料庫表：`OCRD`, `OCPR`, `OOCR`, `OCRN`, `ORTT`
- `CommUtil.SendMail()` 郵件發送函式

---

## ✅ 功能檢查清單

### 基本功能
- [x] 供應商選擇與聯絡人連動
- [x] 收貨地址、日期、幣別選擇
- [x] 匯率自動更新（從 SAP B1）
- [x] 附件上傳/下載（檔名清理 + 時間戳記）

### 費用明細
- [x] 新增/刪除明細行
- [x] 費用類別與總帳科目連動
- [x] 自動計算金額（數量 × 單價）
- [x] 自動計算本幣金額（金額 × 匯率）
- [x] 合計顯示

### MDR 發票明細
- [x] 表頭資訊即時同步
- [x] 新增/刪除發票明細
- [x] 自動計算發票總額（未稅 + 稅額）
- [x] 金額總和驗證（與 AP 比對）

### 審核功能
- [x] 權限控制（CanApproveExpense）
- [x] 審核意見輸入
- [x] 放行/駁回功能
- [x] 郵件通知（使用 CommUtil）

### 資料驗證
- [x] 必填欄位驗證
- [x] 數量、金額正數驗證
- [x] 明細至少 1 筆驗證
- [x] MDR 金額驗證

---

## 📞 需要協助？

請參閱 `整合說明_ExpenseClaimForm.md` 中的：
- 📖 完整整合步驟
- 🧪 測試流程
- 🐛 常見問題排查
- 💡 後續開發建議

---

## 📝 版本記錄

| 版本 | 日期 | 說明 |
|------|------|------|
| v1.0 | 2025-11-05 | 初始版本，完整功能實作 |

---

**提醒**：整合前請務必備份資料庫和現有程式碼！

**所有檔案已產生完成，祝開發順利！** 🎉
