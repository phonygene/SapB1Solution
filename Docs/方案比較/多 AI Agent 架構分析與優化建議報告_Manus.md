# 多 AI Agent 架構分析與優化建議報告

**作者**：Manus AI
**日期**：2026年01月02日
**專案目標**：費用申請單（ExpenseClaim）功能開發與登入/首頁風格優化
**參考文件**：使用者提供的《MULTI_AGENT_ARCHITECTURE_RFC.md》

---

## 1. 總體架構評估與契合度分析

使用者提出的多 Agent 協作架構（Manager-Worker 模式）是一個**設計精良、實用性高**的方案，它有效地解決了單一 Agent 串行工作、認知負擔過重以及經驗無法累積等核心問題。特別是其**資訊分層架構**（Layer 1: 全局狀態、Layer 2: 任務交接、Layer 3: 私有工作區）是當前 Agentic AI 領域中，用於平衡 Token 成本與資訊共享的**最佳實踐之一** [1]。

### 1.1 專案需求契合度

| 專案需求 | Agent 架構契合度 | 說明 |
| :--- | :--- | :--- |
| **費用申請單 (ExpenseClaim)** | **高** | 涉及後端邏輯、SAP B1 整合與前端介面，完美契合 **Backend** 和 **UI-UX** Agent 的協作流程。Manager 可將任務拆分為 API 實作（Backend）與介面呈現（UI-UX），並透過 Layer 2 進行交接。 |
| **登入/首頁風格優化** | **高** | 這是 **UI-UX** Agent 的核心職責。它可以獨立或與 **QA** Agent 協作，專注於 ASP.NET Web Forms 的樣式調整、響應式設計與主題系統導入。 |
| **SAP B1 系統整合** | **極高** | 透過搜尋發現，已有針對 SAP B1 的 **Model Context Protocol (MCP) Server** 實踐 [2]。這使得 Backend Agent 可以透過 MCP 工具直接與 SAP Service Layer 互動，而非傳統的複雜 SDK 呼叫，極大地簡化了舊系統整合的難度。 |
| **舊系統代碼 (VB.NET/ASP.NET)** | **中** | 雖然 Agent 可以處理舊代碼，但 ASP.NET Web Forms 的生命週期與事件模型複雜，AI 處理時的**錯誤率可能較高**。這需要 QA Agent 投入更多資源進行代碼審查與測試。 |

### 1.2 架構優勢與挑戰

| 項目 | 評估 | 說明 |
| :--- | :--- | :--- |
| **資訊分層** | **優勢** | 有效控制 Token 成本，確保 Agent 僅讀取完成任務所需的資訊，避免資訊過載。 |
| **衝突檢測** | **優勢** | `fileConflicts` 機制將多 Agent 協作視為**分散式系統問題**處理 [3]，是並行開發的必要保障。 |
| **經驗累積** | **優勢** | `work-logs/` 的設計為未來的 Agent **自我優化（Self-Correction）**提供了數據基礎。 |
| **Manager 瓶頸** | **挑戰** | 所有的任務拆分、協調、衝突解決都集中在 Manager，可能導致複雜任務的**響應延遲**。 |
| **狀態同步** | **挑戰** | 依賴 Agent 主動更新 `active-tasks.json`，若 Agent 失敗或未及時更新，可能導致其他 Agent 讀取到**過時狀態**。 |

## 2. 優化建議與更佳實踐

雖然現有架構已非常穩健，但結合 2026 年最新的 Agentic AI 趨勢，建議進行以下優化，以提高效率和發展性：

### 2.1 引入「混合式協作模式」（Hybrid Collaboration）

使用者目前傾向於 Manager-Worker 的**中心化**模式。建議在 Manager-Worker 模式的基礎上，引入**去中心化**的「Swarm Intelligence」（蜂群智慧）概念，形成**混合式協作模式** [4]。

| 模式 | 適用場景 | 實踐建議 |
| :--- | :--- | :--- |
| **Manager-Worker** | **複雜、跨領域任務**（如費用申請單從後端到前端的完整開發）。 | 維持現有架構，Manager 負責拆分任務、建立 Layer 2 交接文件。 |
| **Peer-to-Peer (Swarm)** | **單一領域的細節任務**（如 UI-UX Agent 發現樣式問題，直接與 QA Agent 協商解決）。 | 允許 Agent 在 Layer 3 發現問題時，直接在 `handoff/{task-id}/` 下創建一個**臨時協商文件**（如 `negotiation.md`），並標註對方 Agent。Manager 僅需監控此類文件的創建，無需介入細節。 |

### 2.2 強化狀態同步與衝突解決機制

#### 2.2.1 狀態同步：從「拉取」到「推送」

Agent 依賴讀取 `active-tasks.json` 來判斷任務狀態（拉取模式）。建議 Manager 應扮演一個**輕量級通知服務**的角色，實現「推送」機制：

1.  **Manager 監控**：Manager 應持續監控 Layer 2 的 `handoff/{task-id}/output.md` 文件。
2.  **狀態變更觸發**：一旦偵測到 `output.md` 寫入完成，Manager 應立即更新 `active-tasks.json`，並**主動**在目標 Agent 的 Layer 3 工作區 (`workspace/{agent}/current.md`) 中**寫入通知**，告知其依賴的任務已完成。
3.  **實作技術**：這可以透過一個簡單的**檔案系統監聽服務**（如 `inotify` 或 Node.js 的 `fs.watch`）來實現，避免 Agent 頻繁輪詢，降低 Token 成本。

#### 2.2.2 衝突解決：引入 AI 輔助的合併（AI-Assisted Merging）

現有的衝突檢測機制（鎖定檔案）是必要的，但過於保守。當發生衝突時，Agent 應嘗試**自動解決**，而非直接阻塞：

1.  **衝突發生**：Agent 發現目標檔案被鎖定。
2.  **嘗試合併**：Agent 不進入 `blocked` 狀態，而是將其修改與鎖定 Agent 的修改進行**三方合併**（Three-way Merge），並將合併結果與衝突報告寫入 Layer 3 工作區。
3.  **Manager 審核**：Manager 偵測到合併報告後，可呼叫一個專門的 **Merge Agent** 進行審核，若合併成功，則自動提交；若失敗，才將任務設為 `blocked`，並通知使用者介入。

### 2.3 舊系統現代化與 MCP 實踐

針對 VB.NET/ASP.NET Web Forms 的舊系統，應將 Agent 的工作重點放在**現代化**上：

1.  **Backend Agent**：應將業務邏輯從 `.aspx.vb` 檔案中分離出來，重構成獨立的 **.NET Core Web API** 服務。這使得 Agent 可以專注於更現代的 C# 代碼，減少對 Web Forms 生命週期的依賴。
2.  **UI-UX Agent**：應將精力放在**漸進式現代化**（Progressive Modernization）上，例如使用 **Blazor** 或 **React/Vue** 框架來重寫費用申請單的特定組件，並透過 Web Forms 的互操作性嵌入。
3.  **SAP B1 MCP 整合**：
    *   **Backend Agent** 應將 MCP Server 視為一個**工具（Tool）**，而非代碼庫的一部分。
    *   Agent 應學習如何使用 MCP Server 提供的工具集（如 `Session Tools`, `Access to Service Layer Objects`）來執行 CRUD 操作，而不是直接操作 SAP B1 的舊式 SDK。

## 3. 實踐方案總結

使用者提出的架構**非常適合**作為開發 Agentic AI 系統的起點。建議的實踐方案是：

1.  **實作技術**：傾向於使用**多個終端機同時運行**（選項 A），因為這能最簡單地實現並行，且每個 Agent 擁有獨立的環境和狀態，符合分散式系統的原則。
2.  **Manager 角色**：採用**混合模式**（選項 C），但應將 Manager 的職責從「執行者」轉變為「**協調者與通知者**」，專注於狀態同步、衝突解決與任務拆分。
3.  **發展性**：此架構具備極高的發展性。一旦 Manager 角色過重，可以逐步將其部分職責（如衝突解決、任務拆分）下放給專門的 **Meta-Agent**，最終過渡到更接近 **LangGraph** 或 **CrewAI** 所倡導的**工作流程圖（Workflow Graph）**模式 [5]。

總結來說，您的架構設計**邏輯清晰、考量周全**，特別是在資訊分層與衝突檢測方面，已站在當前技術的前沿。只需在狀態同步和衝突解決機制上進行微調，並結合 MCP Server 進行舊系統整合，即可構建一個**高效、穩定且具備高度發展性**的 AI 協作開發環境。

---

## 參考資料

[1] **Designing Multi-Agent Intelligence** - Microsoft for Developers. (2025). *強調避免資訊過載，並建議 Agent 知識領域與行動範圍不應過度重疊。*
[2] **MCP Server for SAP Business One** - CompuTec Learn. (2026). *詳細說明如何透過 CompuTec AppEngine 3.0 啟用 MCP Server，讓 Claude Code 等 AI 客戶端能連接 SAP B1。*
[3] **When AI Tools Fight Each Other: The Hidden Chaos of Multi-Agent Workflows** - Medium. (2025). *指出多 Agent 工作流應被視為分散式系統問題，需要明確的協調協議和共享狀態管理。*
[4] **Multi-Agent collaboration patterns with Strands** - AWS Blogs. (2025). *介紹 Swarm（蜂群）模式，其中 Agent 直接交換資訊，適用於需要高頻率、去中心化協作的場景。*
[5] **Building Multi-Agent Systems: Hands-On Experience** - AltexSoft. (2025). *建議在設計多 Agent 系統時，應從定義問題和目標開始，選擇合適的架構（如 Leader-Worker 或 Graph-based）。*
