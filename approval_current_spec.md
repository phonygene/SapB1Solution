# SapB1Solution 簽核系統「目前規格」工作稿
日期：2026-02-05

本文件用來記錄我們目前討論到的規格與共識，作為重啟 CLI 後的持續依據。

## 1. 目標與範圍
- 以同一平台內完成「費用申請單 / 請購單」的簽核整合。
- 公司規模約 300 人、跨多公司/多據點，簽核層級不多。
- 希望簡潔、可落地、維護成本低，不使用大型 BPMN 引擎。
- 核心想法：以平台唯一 ID `jID` 為流程主鍵，簽核系統依 `jID` 拉單據資料，避免狀態與資料發散。

## 2. 重要名詞釐清
- `sftype`：流程種類，只有 3 種（1/2/3）。
- `sfid`：表單/流程定義的 ID（很多個）。每個 sfid 對應一個 sftype。

### 2.1 sftype 三種流程類型（現有系統）
- sftype=1：固定流程（部門層級 + 表單固定簽核人 + 歸檔 + 知悉）。
- sftype=2：部門層級 + 可補人（可加簽/知悉，但不應改掉原本鏈）。
- sftype=3：完全自訂（送審者決定簽核/歸檔/知悉）。

### 2.2 @XSPWT / @XSDET（現有系統）
- `@XSPWT`：每張單據的簽核人員清單（含簽核/歸檔/知悉）。
- `@XSDET`：部門排除（sfid + uid + deptcode）。用來排除固定簽核人或歸檔人。

## 3. 目前系統現況（摘要）
- 存在兩條簽核路徑：
  - SAP B1 UDT 簽核引擎（@XSFTT/@XASCH/@XSPWT...）
  - jOPCH/jOPRQ 簡化簽核（ApprovalStatus=W/A/R 直接寫單據表）
- UDT 引擎功能完整但複雜，資料在 SAP DB。
- jOPCH/jOPRQ 簡化簽核與 UDT 引擎未整合。
- `U_PID` 目前只是 UI 欄位與查詢條件，非真正流程主鍵。
- `jID` 來自 OJID，為全平台單據唯一主鍵。

## 4. 目前已確認/推論的關鍵規則
- 部門層級（signlevel）最大約 3 層（遇 topsignoffs 或 NA 會停止）。
- sftype=2：可補第二/第三簽核角色，但不應改掉原本系統帶入的人；
  例外僅見 sfid=101。
- sftype=3：完全自訂，內控風險較高，應視為低風險或特例流程。

## 5. 期望的新簽核引擎方向（共識草案）
- 以 `jID` 為唯一流程主鍵，簽核表與單據表解耦。
- 單據表只保留一個狀態欄位，流程細節存在簽核表。
- 保持少層級、可配置、易維護。

### 5.1 建議最小資料結構（草案）
- ApprovalInstance
  - approval_id (PK), jID, doc_type, company_id
  - status (DRAFT/PENDING/APPROVED/REJECTED/CANCELED)
  - current_step, created_by, created_at, updated_at, revision

- ApprovalStep
  - approval_id, step_no, mode (SEQUENTIAL/PARALLEL), min_required

- ApprovalAssignee
  - approval_id, step_no, assignee_type (USER/ROLE/TEAM)
  - assignee_id, status, acted_at, comment

- ApprovalAction（不可變更歷程）
  - approval_id, actor_id, action
  - from_status, to_status, snapshot_json/hash, created_at

- ApprovalRule
  - doc_type, company_id, amount_min/max, dept, project, route_json

### 5.2 狀態最小集合（草案）
- DRAFT -> PENDING -> APPROVED / REJECTED -> COMPLETED
- CANCELED

## 6. 目前的實作/相容性要求
- 保留 SAP B1 放行後寫入流程（B1PostStatus 控制）。
- 需要完整歷程（審核人、時間、動作）。
- 代理簽核功能需保留（可先簡化為代理設定表）。

## 7. 開放問題 / 待確認
- 是否要保留「歸檔/知悉」為正式關卡，或轉為通知即可。
- `U_PID` 是否要保留為對外編號，或改由 approval_id 取代。
- 權限模型是否維持 Approver/AP_App/PU_App，或逐步轉成角色制。

---
本文件為工作稿，後續將持續更新。
