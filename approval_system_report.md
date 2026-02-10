# SapB1Solution 簽核系統現況與建議（依 jID 整合）
日期：2026-02-04
版本：v1.1

## 0. 目的與範圍
本文件整理 SapB1Solution 目前的簽核機制、資料結構與流程行為，並提出「以 jID 為核心」的精簡可落地方案，最後做出與既有方案的優劣比較。

## 1. 資料來源與方法
- 資料庫（讀取結構）：
- jtdb（.\SQLEXPRESS2008R2）
- SAP B1 DB：JTTST（SAP-NEW-TST）
- 主要程式碼：
- MgmSP\commcode\CommSignOff.vb
- MgmSP\signoff\cLsignoff.aspx.vb
- MgmSP\ExpenseClaimForm.aspx.vb
- MgmSP\PurchaseRequestForm.aspx.vb
- 備註：僅讀取結構與程式邏輯，未進行任何寫入操作。

## 2. 現況總結（最短版）
- 系統內同時存在兩條簽核路徑。
- SAP B1 UDT 簽核引擎：@XSFTT/@XASCH/@XSPWT/... 功能完整但複雜，核心資料在 SAP DB。
- jOPCH/jOPRQ 簡化簽核：ApprovalStatus (W/A/R) 直接寫單據表，與 UDT 引擎未整合。
- U_PID 在 jOPCH/jOPRQ 目前是 UI 欄位與查詢條件，並非真正簽核流程的主鍵。
- 若以 jID 作為單一主鍵整合，可大幅降低狀態與資料發散問題。

## 3. jtdb（平台資料）結構整理
### 3.1 jOPCH（費用申請單主檔）
關鍵欄位：
- jID (int, PK)
- ApprovalStatus (nvarchar(20), 預設 Pending)
- ApprovedBy / ApprovedDate / ApprovalComments / ApprovalDate / ApprovalTime
- U_PID (int, 簽核 PID，UI 顯示/查詢用)
- B1PostStatus / B1PostDate / B1ErrMsg
- DocTotal / DocCurrency / DocDate / DocDueDate 等

觀察：
- 程式實際用 W/A/R（待審/核准/駁回），與資料庫預設 Pending 不一致。
- U_PID 在資料表是 int，但 DB_Create_JET_Tables.sql 中定義為 NVARCHAR(50)，存在規格不一致。

### 3.2 jOPRQ（請購單主檔）
關鍵欄位：
- jID (int, PK)
- ApprovalStatus (nvarchar(20), 預設 Pending)
- ApprovedBy / ApprovedDate / ApprovalComments
- U_PID (int)
- B1PostStatus / B1PostDate / B1ErrMsg
- DocTotal / DocDate / ReqDate 等

觀察：
- Insert 時直接設 W（待審）。
- 仍存在 Pending/A/R/W 混用的歷史包袱。

### 3.3 jPCH1 / jPRQ1（明細表）
- 以 jID + LineNum 作為明細鍵。
- 與簽核邏輯關係不大，屬單據內容。

### 3.4 OJID（全域 jID 來源）
- jID 為 PK，並記錄 DocType/jUser。
- jID 在所有單據間具有全域唯一性，可作為簽核引擎主鍵。

### 3.5 User（權限欄位）
- Approver / AP_App / PU_App 等欄位作為審核權限。
- signlevel / signprice 存在，但目前只在 UDT 引擎邏輯中使用。

## 4. SAP B1 UDT 簽核引擎（現有完整版）
### 4.1 主要資料表
- @XSFTT：表單種類定義（sfid/sfname/sftype）
- @XASCH：簽核主檔（docnum/status/sid/subject/sfid/price/area/dept）
- @XSPWT：簽核人員/關卡清單（docentry/uid/seq/signprop/status/signdate/comment/receivedate）
- @XSPMT：內定簽核人設定（sfid/uid/seq/prop）
- @XSPHT：簽核歷程（docentry/uid/signdate/status/comment/flowseq）
- @XSMLS：表單內容與附加資訊（docentry/itemcode/quantity/price/head/descrip）
- @XSDET：部門排除/部門簽核設定（sfid/uid/deptcode）
- @XSTDT：追蹤/待辦（docentry/status/subject/incharge/traceperson）
- @XSPAT：自訂簽核人組（uid/ownid/signpid/signpname/prop）

### 4.2 @XASCH.status（推論語意）
- E / D：未送審（草稿/未送出）
- O：簽核中
- F：簽核完成（結案）
- T：歸檔完成
- B：退回
- R：抽回
- C：作廢

### 4.3 @XSPWT.status（推論語意）
- 0：未到/備核
- 1：待簽核
- 2：核准
- 3：反對/駁回
- 5：取消
- 10：跳過簽核
- 100：重新送審/送出標記
- 103：歸檔完成（signprop=1）
- 104：已知悉（signprop=2）

### 4.4 signprop（簽核角色）
- 0：簽核
- 1：歸檔
- 2：知悉

### 4.5 核心流程（高度摘要）
- 送審：產生 @XSPWT；@XASCH.status=O；首關 status=1，後續 0。
- 核准：當前關 status=2；若最後關 -> @XASCH.status=F；啟動歸檔與知悉。
- 駁回：
- 若 innerloop=1 或最後關 或選擇退回送審者 -> 全部清 0，送審者 status=1，@XASCH.status=B。
- 否則：當前關 status=3，但仍推進下一關（反對但不阻擋）。
- 抽回：送審者可抽回，@XASCH.status=R。
- 跳過：管理者 skip -> status=10，推進下一關。
- 作廢：@XASCH.status=C；送審者 status=5。
- 歸檔：signprop=1 完成，@XASCH.status=T。
- 知悉：signprop=2 完成，status=104。

## 5. jOPCH/jOPRQ 目前簽核行為（簡化版）
### 5.1 費用申請單（jOPCH）
- 單一關卡核准/駁回，ApprovalStatus=W/A/R。
- 有樂觀鎖定與 B1PostStatus 防重入。
- A 時觸發 SAP B1 AP Invoice 建立。
- 有 AuditLogger 記錄狀態變更（jOPCH）。

### 5.2 請購單（jOPRQ）
- 單一關卡核准/駁回（PU_App 權限）。
- 無明確的狀態轉換驗證與樂觀鎖定（直接 UPDATE）。
- 無獨立歷程表。

## 6. 主要落差與風險
- 簽核引擎在 SAP DB，費用/請購在 jtdb，兩條流程未同步。
- jOPCH/jOPRQ 無完整簽核歷程；Audit 需求只能靠 jOPCH 的 AuditLogger（jOPRQ 無）。
- U_PID 被標示為「簽核 PID」，但目前未連結任何簽核引擎。
- ApprovalStatus 型別與值不一致（Pending/W/A/R）。
- 請購單缺乏樂觀鎖定與狀態轉換驗證，存在競爭條件風險。

## 7. 建議的 jID 整合方案（精簡可落地）
### 7.1 核心思路
- 以 jID 作為唯一流程主鍵。
- 單據資料仍在 jOPCH/jOPRQ；簽核流程與歷程獨立表管理。
- 單據表只保留一個狀態欄位（ApprovalStatus），避免狀態發散。

### 7.2 建議最小資料結構
- ApprovalInstance
- approval_id (PK)
- jID, doc_type, company_id
- status (DRAFT/PENDING/APPROVED/REJECTED/CANCELED)
- current_step, created_by, created_at, updated_at
- revision (int，支援重送)

- ApprovalStep
- approval_id, step_no, mode (SEQUENTIAL/PARALLEL)
- min_required

- ApprovalAssignee
- approval_id, step_no, assignee_type (USER/ROLE/TEAM)
- assignee_id, status, acted_at, comment

- ApprovalAction（不可變更歷程）
- approval_id, actor_id, action
- from_status, to_status
- snapshot_json/hash, created_at

- ApprovalRule（規則路由）
- doc_type, company_id, amount_min, amount_max, dept, project
- route_json

### 7.3 狀態最小集合
- DRAFT -> PENDING -> APPROVED / REJECTED -> COMPLETED
- CANCELED（作廢）

### 7.4 與 jOPCH/jOPRQ 的整合方式
- 送審時建立 ApprovalInstance，將 ApprovalStatus 更新為 W。
- 每次簽核行為寫入 ApprovalAction。
- 只有引擎能更新 jOPCH/jOPRQ.ApprovalStatus。
- 放行後觸發 SAP API（保持現有流程）。

## 8. 方案對照（既有引擎 vs 建議方案）
### 8.1 既有 UDT 引擎
優點：
- 功能完整：代理、歸檔、知悉、催簽、PDF 簽條
- 已有表結構與流程定義

缺點：
- 與 jOPCH/jOPRQ 分離，難以以 jID 統一
- 狀態與語意複雜，維護成本高
- 依賴 WebForms 舊程式碼與 SAP DB UDT

### 8.2 建議 jID 引擎
優點：
- 完全整合 jOPCH/jOPRQ
- 狀態簡化、易維護
- 適合 300 人規模與少層級簽核

缺點：
- 需重新實作部分功能（代理、歸檔、知悉）

### 8.3 最佳折衷
- 新建 jID 引擎，優先支援費用/請購。
- 舊 UDT 引擎留作其他表單或逐步淘汰。

## 9. 建議落地順序
1. 統一 ApprovalStatus 語意（W/A/R）。
2. 建立最小簽核引擎表（ApprovalInstance/Step/Assignee/Action）。
3. 在費用/請購送審時寫入引擎。
4. 在簽核介面只讀 jID 對應單據資料。
5. 完成放行後 SAP API 呼叫保持原流程。
6. 補上代理/催簽（可後置）。

## 10. jID 與 UDT docnum 的核心差異
- UDT 引擎使用 docnum/docentry（SAP DB）作為流程主鍵。
- jOPCH/jOPRQ 使用 jID（平台 DB）作為單據主鍵。
- 目前兩者沒有直接映射欄位；U_PID 只是 UI 欄位，非流程外鍵。
- 若要整合 UDT，引擎必須持有 jID 或另建 mapping 表（維護成本高）。

## 11. 權限模型差異
- jOPCH：Approver / AP_App
- jOPRQ：PU_App
- UDT：signlevel / signprice + 部門/層級規則

建議：
- 在新引擎保留現有 Approver/AP_App/PU_App 作為起點。
- 之後逐步改為「角色 + 規則路由」。

## 12. 狀態語意對照（建議映射）
- @XASCH.E/D -> DRAFT
- @XASCH.O -> PENDING
- @XASCH.F -> APPROVED
- @XASCH.T -> ARCHIVED (可視為 COMPLETED)
- @XASCH.B -> REJECTED
- @XASCH.R -> WITHDRAWN
- @XASCH.C -> CANCELED

- jOPCH/jOPRQ.W -> PENDING
- jOPCH/jOPRQ.A -> APPROVED
- jOPCH/jOPRQ.R -> REJECTED
- Pending -> DRAFT 或 PENDING（需統一）

## 13. 需要保留的「進階功能」建議（可選）
- 代理簽核：可先用「臨時代理人」配置表替代。
- 歸檔/知悉：若必要，可在 APPROVED 後自動派發通知，不一定做成關卡。
- 催簽：先用排程 + 待辦查詢發送提醒即可。
- 反對但不阻擋（innerloop）：可先不支援，若未來需求再擴充。

## 14. 立即可執行的風險修補
- jOPRQ 增加狀態轉換驗證與樂觀鎖定（避免同時核准/駁回）。
- 統一 ApprovalStatus 值域，清理 Pending/Approved/Rejected 舊值。
- 將 U_PID 重新命名或明確說明為「專案/簽核編號」，避免誤解。

---
本文為 v1.1 版本整理。下一步可繼續輸出：
- 狀態轉換圖
- API 設計草案
- DB migration 設計
- UDT 引擎與 jID 引擎逐欄位對照
