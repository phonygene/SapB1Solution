## 工作流程規範 (Workflow Standards)

### Shopfloor 協作模式

本專案採用 **Shopfloor 協作模式**，目的是優化 Claude 與使用者的協作效率。

---

## 核心原則

### 1. 檔案輸出規範

#### Shopfloor 流程適用範圍（僅限 SQL）

**只有 SQL 資料庫操作**需要先輸出到 `shopfloor/Claude_TMP/`：
- CREATE TABLE / ALTER TABLE / DROP TABLE
- 資料遷移腳本
- 批次 INSERT / UPDATE / DELETE
- 複雜查詢腳本

```
shopfloor/Claude_TMP/
└── SqlQuery/       - SQL 腳本（建表、查詢、修改）
```

#### SQL 檔案命名規範
- 格式：`[編號]_[用途]_[物件名稱].sql`
- 範例：`01_CreateTable_user_vendor_taxid.sql`

#### 程式碼改動（不使用 Shopfloor）

**所有程式碼改動改用 Git 版本控制**，不再輸出到 shopfloor：

1. **改動前**：確認目前 git 狀態乾淨，或先 commit 現有變更
2. **改動中**：直接修改專案檔案
3. **改動後**：立即 commit，方便追蹤和回復

#### Git 流程優點
- 可隨時 `git diff` 查看改動內容
- 可隨時 `git restore` 回復錯誤修改
- 完整的變更歷史記錄
- 比 shopfloor 流程更高效

---

### 2. 程式碼顯示規範

**不在對話中貼大段程式碼**

#### 規則
- ❌ **禁止**：在對話中貼超過 **20 行**的程式碼
- ✅ **允許**：簡短的程式碼片段（少於 10 行）用於說明
- ✅ **允許**：錯誤訊息和 Log 輸出

#### 優點
- 節省 Token（不在對話中顯示大量代碼）
- 減少對話視覺疲勞
- 方便使用者直接使用（複製貼上）
- 保留完整歷史（所有產出都有檔案記錄）
- 易於版本控制（可 git add 這些檔案）

---

### 3. 檔案說明規範

**每個輸出檔案開頭必須加上使用說明**

#### 部分代碼 TXT 檔案格式
```
=====================================================================
  [檔案名稱] - [用途說明]
=====================================================================
更新時間：YYYY-MM-DD HH:mm

一、用途說明
---------------------------------------------------------------------
[詳細說明這個檔案的用途]

二、目標位置
---------------------------------------------------------------------
應加入到：[目標檔案路徑]
位置：[應該放在哪個區域或函式中]

三、使用方式
---------------------------------------------------------------------
1. [步驟1]
2. [步驟2]
3. [步驟3]

四、需要檢查的點
---------------------------------------------------------------------
- [ ] [檢查項目1]
- [ ] [檢查項目2]

五、完整代碼
---------------------------------------------------------------------
[程式碼內容]
=====================================================================
```

---

## 工作流程

### 標準流程

```
1. Claude 產生檔案到 shopfloor/Claude_TMP/
   ↓
2. Claude 在對話中簡要說明：
   - 本次產生的檔案清單
   - 使用順序
   - 注意事項
   ↓
3. 使用者在 VS Code/Visual Studio 開啟檔案
   ↓
4. 使用者根據檔案開頭的說明進行操作：
   - SQL 腳本 → 在 SSMS 執行
   - 完整檔案 → 加入專案
   - 部分代碼 → 複製到目標位置
   ↓
5. 使用者簡短回報結果：
   - "執行成功" / "建表完成"
   - "檔案已加入專案"
   - "報錯：[錯誤訊息]"
   - "需要調整：[具體需求]"
   ↓
6. Claude 根據回報繼續產生下一個檔案
```

---

## 程式碼改動流程

### Git 版本控制流程

**所有程式碼改動**（包括新功能、Bug 修正、重構）都使用 Git：

```
1. 改動前：git status 確認狀態
   ↓
2. 如有未提交變更：先 commit 或 stash
   ↓
3. 直接修改專案檔案
   ↓
4. 改動後：git diff 確認變更內容
   ↓
5. 詢問用戶是否 commit
```

### 回復錯誤修改

如果改動有誤，可使用：
- `git restore <file>` - 回復單一檔案
- `git restore .` - 回復所有未提交變更
- `git revert <commit>` - 回復已提交的 commit

### SQL 資料庫操作（仍使用 Shopfloor）

**只有 SQL 操作**需要先輸出到 `shopfloor/Claude_TMP/SqlQuery/`：
- 因為資料庫變更難以回復
- 需要用戶在 SSMS 手動執行確認

---

## 改動前檢查清單

### 程式碼改動檢查

在修改程式碼前，先確認：

- [ ] git status 是否乾淨？（如有未提交變更，先處理）
- [ ] 是否了解用戶的需求？（不確定就先問）
- [ ] 新增檔案是否需要 UTF-8 BOM？
- [ ] 新增 .aspx 是否需要更新 .vbproj？

### SQL 操作檢查

在執行 SQL 前，先確認：

- [ ] 是否已輸出到 `shopfloor/Claude_TMP/SqlQuery/`？
- [ ] 是否已向用戶說明 SQL 內容？
- [ ] 用戶是否已確認執行？

---

## 溝通方式

### Claude 的回應格式

當產生檔案時，使用以下格式回應：

```markdown
## 📁 已產生檔案

我已將以下檔案產生到 `shopfloor/Claude_TMP/`：

### SQL 腳本（1 個）
1. `SqlQuery/02_CreateTable_XXX.sql` - [用途說明]

### VB.NET 檔案（2 個）
1. `dNet/XXX.vb` - [用途說明]
2. `dNet/XXX.aspx` - [用途說明]

## 📋 使用順序

1. 先執行 SQL 腳本建立資料表
2. 將 .vb 檔案加入專案的 commcode/ 目錄
3. 將 .aspx 檔案加入專案根目錄

## ⚠️ 注意事項

- [注意事項1]
- [注意事項2]

請執行後回報結果，我會根據您的回報繼續下一步。
```

### 使用者的回應格式

使用者只需簡短回報：

- ✅ "執行成功"
- ✅ "已加入專案"
- ✅ "建表完成"
- ❌ "報錯：[錯誤訊息]"
- 🔄 "需要調整：[具體需求]"

---

## 命名規範與檢查

### SQL 保留關鍵字檢查規則

**強制執行規則**：當使用者提出任何命名（欄位、變數、表名等）時，必須主動檢查是否與 SQL Server 保留關鍵字衝突。

#### 檢查時機

1. **資料表設計時**
   - 表名稱
   - 欄位名稱
   - 索引名稱

2. **程式碼撰寫時**
   - VB.NET 變數名稱（涉及 SQL 查詢）
   - 參數名稱
   - 函式名稱（涉及資料庫操作）

3. **使用者提議命名時**
   - 任何可能與資料庫互動的命名

#### 提醒格式

當發現關鍵字衝突時，使用以下格式提醒：

```markdown
⚠️ **命名衝突警告**

您提議的名稱 `[NAME]` 是 SQL Server 保留關鍵字。

**問題說明**：
- 用途：[說明這個關鍵字在 SQL 中的用途]
- 影響：會導致查詢時需要使用方括號 `[NAME]` 包裹

**建議替代方案**：
1. `[建議1]` - [說明]
2. `[建議2]` - [說明]
3. `[建議3]` - [說明]

建議使用哪一個？或您有其他想法？
```

#### 常見保留關鍵字清單

**高風險關鍵字**（容易誤用）：
- `ON`, `ORDER`, `USER`, `TABLE`, `INDEX`, `KEY`
- `DATE`, `TIME`, `TIMESTAMP`, `VALUE`, `VALUES`
- `CHECK`, `DEFAULT`, `LEVEL`, `PERCENT`
- `TYPE`, `OPTION`, `ROLE`, `LOGIN`

**資料操作類**：
- `SELECT`, `INSERT`, `UPDATE`, `DELETE`
- `FROM`, `WHERE`, `JOIN`, `GROUP`, `HAVING`
- `UNION`, `EXCEPT`, `INTERSECT`

**資料定義類**：
- `CREATE`, `ALTER`, `DROP`
- `PRIMARY`, `FOREIGN`, `UNIQUE`
- `DATABASE`, `SCHEMA`, `VIEW`

**短詞關鍵字**：
- `IN`, `AS`, `IS`, `OR`, `AND`, `NOT`
- `GO`, `IF`, `OF`, `TO`, `BY`

#### 建議的命名規範

**遇到關鍵字衝突時的解決方案**：

1. **加上後綴**
   - `Order` → `OrderNo`, `OrderID`
   - `User` → `UserID`, `UserName`
   - `Date` → `OrderDate`, `CreateDate`

2. **使用完整描述**
   - `Type` → `ProductType`, `DocumentType`
   - `Level` → `AccessLevel`, `PriorityLevel`

3. **改用同義詞**
   - `Status` → `State`
   - `Type` → `Category`
   - `Order` → `Sequence`

4. **加上業務前綴**
   - `Table` → `DataTable`, `MappingTable`
   - `Index` → `SortIndex`, `DisplayIndex`

#### 承諾

- ✅ 每次使用者提出命名時，主動檢查是否為保留關鍵字
- ✅ 發現衝突時立即提醒，不等使用者發現問題
- ✅ 提供至少 3 個具體的替代方案
- ✅ 說明為什麼這個命名會有問題

**記錄日期**：2025-11-20
**提出者**：Jason
**執行者**：Claude

---

## 參考文件

- **詳細說明**：`shopfloor/Claude_TMP/etc/README_協作模式說明.txt`
- **Session 管理**：`agent-os/standards/global/session-management.md`
- **溝通規範**：`agent-os/standards/global/communication-standards.md`
- **SQL 保留關鍵字完整清單**：https://learn.microsoft.com/en-us/sql/t-sql/language-elements/reserved-keywords-transact-sql

---

**最後更新**：2025-11-20
**維護者**：Claude + Jason
