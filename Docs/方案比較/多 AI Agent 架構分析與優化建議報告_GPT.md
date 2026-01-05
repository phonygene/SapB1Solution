# 多 AI Agent 架構適配性評估（依你的 RFC + 近期趨勢整理）

產出時間：2026-01-02 15:33（台北）

> 依據你提供的《多 Agent 協作架構設計文件（RFC v0.1, 2026-01-02）》fileciteturn1file3L1-L6，以及近期多 Agent / Coding Agent 的實務趨勢與框架能力整理（LangGraph / AutoGen / n8n 等）citeturn1search2turn1search3turn1search26turn1news44。

---

## 0. 我能/不能看你的 GitHub 專案

- **我目前沒有拿到你的 GitHub Repo 連結**，因此本文件是**以你 RFC 描述的技術棧（VB.NET / ASP.NET Web Forms、SAP B1 整合）與你提到的改動範圍（ExpenseClaim + 登入/首頁樣式優化）**為主來做適配性評估fileciteturn1file3L14-L20。  
- 若你之後提供 Repo URL（且為 public 或你允許我存取的來源），我可以再把「具體檔案耦合度、實際改動點、可切分任務邊界」做更精準的二次評估。

---

## 1. 你的架構方向對不對？

### 1.1 方向是對的：你抓到 3 個關鍵痛點
你希望解決的痛點是：
- 串行等待、認知負擔、重複犯錯、以及多 AI 之間狀態無法共享fileciteturn1file3L22-L29  
- 目標是並行、能交接、能累積經驗、讓你只在關鍵決策介入fileciteturn1file2L1-L6  

這些都符合近一年「coding agents」落地時的共識：**最痛的不是模型不會寫，而是流程、上下文、驗證、可控性**。例如近期針對 2026 工作流的建議，也強調「要提供足夠上下文與約束，否則品質不穩定」citeturn1search0。

### 1.2 你採用「資訊分層」是很好的決策
你 RFC 用 Layer 1/2/3 讓不同 Agent 只讀必要資訊，避免 token 爆炸與資訊過載，是實務上很有效的折衷（也能避免 Manager 成為瓶頸）fileciteturn1file0L20-L22。  
在「需要可控性」這件事上，像 LangGraph 這類框架也把重點放在 **狀態機 + human-in-the-loop + guard rails**，用流程控制代理不要跑偏citeturn1search2turn1search28。

---

## 2. 但你現在的版本，最可能踩到的坑

### 2.1 對你這種「單人、改既有專案」情境：多 Agent 並行容易被協調成本吃掉
你自己在 RFC 其實已經列出關鍵風險：Manager 會成瓶頸、同步頻率會有成本、以及你明確不想要複雜審批與額外基礎設施fileciteturn1file2L26-L30。  
而你的改動範圍（ExpenseClaim 功能 + 登入/首頁風格）往往具備這些特性：
- Web Forms 常見「UI 與事件處理/資料處理」耦合偏高（同檔或同頁面生命週期綁很緊）
- UI 改動常需要後端同步調整（欄位、驗證、資料回寫）
- 若同時派 UI Agent 與 Backend Agent 併發改同一頁，很容易互相踩到或出現接口不一致

結果是：**你想省時間，但會花更多時間在“對齊與修補”**。

### 2.2 你設計的「檔案鎖」概念可行，但很可能變成摩擦來源
你用 active-tasks.json 的 fileConflicts 來鎖檔，是合理的初步控制fileciteturn1file1L1-L20。  
但在真實開發裡，衝突多半不是「同一檔案」那麼簡單，而是：
- UI/後端約定（欄位名稱、資料型態）不一致
- 同一功能散落多檔：aspx + code-behind + css + js + config
- “看似不同檔”但最後在 runtime 上互相影響

因此只靠檔案鎖會：
- 過度保守：不敢並行，變回串行
- 或過度樂觀：沒鎖到真正的衝突點

### 2.3 「更多 Agent」不是最好的擴張方向：更實際的是「技能庫（skills）」與可重用流程
近期 Anthropic 研究者就提出：與其堆一堆不同 Agent，不如讓一個通用 Agent 搭配一套可組合的「skills（可重用程序知識包）」更務實citeturn1news44。  
這點跟你想要的「經驗累積」其實非常吻合fileciteturn1file2L5-L6：你已經有 work-logs/，下一步更關鍵的是把它產品化成「可反覆套用的技能/清單/模板」。

---

## 3. 我最推薦的「更實際、更有效率」落地版本（保留你想要的精神，但減重）

我建議你把 4 Agent 常駐改成 **2+1** 模式：

### 3.1 常駐 2 個角色：Manager（規格/拆分/整合） + Builder（全端實作）
- **Manager**：你 RFC 定義的任務分派/追蹤/日誌維護保留fileciteturn1file0L13-L18  
- **Builder（取代 Backend + UI-UX）**：在 Web Forms 這種耦合環境下，由同一 Agent 連續改 UI 與 code-behind，通常比拆兩個 Agent 更穩、整合成本更低

### 3.2 第 3 個角色「按需召喚」：QA/Critic（審查 + 測試思維）
- QA 不必常駐並行；在每個子任務完成後，做一次**結構化審查**（可包含：風險點、邊界條件、回歸影響、測試清單）
- 這符合你「不想要過度複雜審批」的限制fileciteturn1file2L26-L30

> 這個 2+1 模式，本質上是在「並行」與「一致性」之間取更適合你的平衡：  
> **把並行用在“模組彼此獨立”時**（例如登入頁 UI 與 ExpenseClaim 後端），  
> **把一致性用在“同頁面/同流程耦合高”時**（ExpenseClaim 同一張表單）。

---

## 4. 你原本架構中，哪些要保留？哪些要改？

### 4.1 建議保留（幾乎不用動）
1. **資訊分層思想**（Layer 1/2/3）fileciteturn1file0L20-L22  
2. **handoff 任務交接資料夾**（這非常適合當“規格/接口契約”載體）fileciteturn1file4L62-L84  
3. **active-tasks.json 最小追蹤欄位**（id / assignee / status / blockedBy / affectedFiles）fileciteturn1file4L15-L28  

### 4.2 建議調整（關鍵）
1. **把 UI-UX 與 Backend 的“常駐分工”改為“依任務耦合度切換”**
   - 耦合高（同頁面）：Builder 一次改完  
   - 耦合低（不同模組）：才拆給 UI-UX/Backend 並行
2. **fileConflicts 改成“提示 + Git 流程優先”**
   - fileConflicts 保留當提示
   - 真正的衝突控制以 Git 分支 / PR / review 為主（你 RFC 也已在用分支前綴）fileciteturn1file0L13-L18
3. **把 work-logs 進化成 skills/**
   - 每個反覆出現的任務（例如：新增欄位、做表單驗證、做 SAP B1 介接、寫 SQL 查詢、做錯誤處理）都整理成：
     - checklist.md
     - prompt.md（給 Builder/QA 的提示模板）
     - examples/（一兩個成功案例）
   - 這比“更多 Agent 互相開會”更能穩定提升產能citeturn1news44

---

## 5. 「最新實踐」你可以直接借來用的幾個要點

### 5.1 以“可控流程”取代“自由對話”
LangGraph / 類似框架把重點放在：**明確狀態、轉移規則、守門（guard）與人類介入點**citeturn1search2turn1search28。  
你不用真的導入 LangGraph，但可以借它的思想：  
- 每個 TASK 都有固定狀態：pending → in_progress → review → completed/failed（你已定義）fileciteturn1file4L20-L28  
- 每次狀態遷移要滿足一個“守門條件”（例如：已更新 spec/handoff、已附測試清單、已跑過最低限度驗證）

### 5.2 把“計畫”顯式化（planning）
多 Agent 系統最怕各做各的，CrewAI 等框架近年也把“planning / roadmap”當核心價值之一citeturn1search5turn1search23。  
你可以把 planning 簡化成：
- spec.md 內一定要有「接口契約 / 影響檔案 / 回歸風險 / 測試點」

### 5.3 觀測性（observability）比更多 Agent 重要
AutoGen 強調分層、可擴充、以及在多代理網路中做開發/除錯citeturn1search3。CrewAI 也在控制台/追蹤上持續強化citeturn1search23。  
對你來說，不一定要上那些平台；但至少做到：
- 每個 TASK 的 input/output 只放在 handoff/{task}/
- 任何 bug 修復都要回寫到 skills/（否則你 RFC 的“避免重複犯錯”會落空）fileciteturn1file3L28-L29

---

## 6. 給你一個最小可行的流程（你可以今天就開始用）

> 下面流程完全符合你 RFC 的精神，但大幅減少“管理負擔”。

1. **Manager：寫 20 行 spec**
   - 目標 / 非目標
   - 需要改的檔案（affectedFiles）
   - 接口契約（欄位/事件/資料格式）
   - DoD（完成判準）+ 最低測試清單
2. **Builder：一次改完（同一上下文）**
   - UI + code-behind + 必要的 SQL/資料處理
   - 每 commit 都帶 task-id
3. **QA/Critic：結構化審查**
   - 風險點（尤其是 SAP B1 整合與資料一致性）
   - 回歸測試建議
4. **work-logs → skills**
   - 若這次踩到坑，把坑寫成一個 checklist 或提示模板

---

## 7. 其他方案（簡述）

- **完整 4 Agent 並行（你 RFC 原案）**：  
  優點：理論上最快；缺點：對單人/高耦合 Web Forms 容易被整合成本吃掉。  
- **導入真正的框架（LangGraph/AutoGen/n8n）**：  
  優點：能做更自動化、更可靠的多步流程citeturn1search2turn1search3turn1search26；缺點：你 RFC 說不想引入額外基礎設施，初期投入可能超過收益fileciteturn1file2L26-L30。

---

## 8. 我對你這個專案的結論（在你目前資訊下）

- 你的 RFC 架構**非常適合當“規格/狀態/交接的骨架”**，尤其是資訊分層與 handoff 機制。fileciteturn1file0L20-L22  
- 但對「ExpenseClaim（Web Forms + SAP 整合）」這種高耦合功能，**常駐拆成 Backend + UI 兩個並行 Agent，收益通常不如成本**。  
- 最實際的落地做法是：**2+1（Manager + Builder + On-demand QA/Critic）**，再把你的 work-logs 進化成 skills/，長期成長會更快、更穩。

---

### 你如果願意再往下一步
把你的 GitHub Repo URL（或直接上傳關鍵檔案：ExpenseClaimForm.aspx / .vb、Login/Home 的頁面與 CSS）丟給我，我可以：
- 用“檔案耦合度”幫你判斷哪些任務可並行、哪些必須串行
- 幫你把 skills/ 的第一批模板（例如：新增欄位、做驗證、串 SAP、寫 SQL 查詢、做例外處理）整理成可直接用的結構
