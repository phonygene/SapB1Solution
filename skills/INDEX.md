# Skills 資源索引

> 🔴 **維護規則**：變更 `skills/` 任何檔案時，必須同步更新此索引

---

## 觸發條件對照表

| 觸發條件 | 資源檔案 | 說明 |
|----------|----------|------|
| **執行 SQL 查詢（MCP 工具）** | `database-guide.md` | 🔴 **必讀** - 選擇正確的資料庫 |
| 建立/修改 Slash Command | `slash-command-standards.md` | 指令格式規範、frontmatter 語法 |
| 後端開發（VB.NET、資料庫） | `backend-checklist.md` | 資料庫操作、金額處理、錯誤處理 |
| 資料存取（ADO.NET、ViewState） | `aspnet-data.md` | 參數化查詢型別、資料儲存位置 |
| UI/前端開發（ASPX、JavaScript） | `ui-checklist.md` | 控制項、CSS、PostBack 處理 |
| SAP B1 整合（Service Layer、DI API） | `sap-checklist.md` | Session 管理、COM 物件釋放 |
| UI 設計（顏色、樣式、元件） | `ui-design-system.md` | 設計規範、禁止事項、允許變更 |
| 所有開發任務 | `general-checklist.md` | Git 規範、檔案編碼、控制項宣告 |
| **工具執行失敗/重試** | `work-logs/insights/tool-errors.md` | 錯誤記錄與已知解決方案 |
| **回復上次工作狀態 / Session 整理** | `session-recovery.md` | 多 Session 整合與最新狀態回復 |

---

## 檔案清單

| 檔案 | 適用 Agent | 更新日期 |
|------|-----------|----------|
| `database-guide.md` | All | 2026-01-15 (新建) |
| `aspnet-data.md` | All (Backend 重點) | 2026-01-13 (新建) |
| `backend-checklist.md` | Backend | 2026-01-07 |
| `general-checklist.md` | All | 2026-01-08 (錯誤處理規範) |
| `sap-checklist.md` | Backend | 2026-01-07 |
| `slash-command-standards.md` | All | 2026-01-07 |
| `ui-checklist.md` | UI-UX | 2026-01-07 |
| `ui-design-system.md` | UI-UX | 2026-01-07 |
| `session-recovery.md` | All | 2026-01-26 (新建) |
| `work-logs/insights/tool-errors.md` | All | 2026-01-08 (新建) |

---

## 協作流程命令

| 命令 | 用途 | 執行者 |
|------|------|--------|
| `/blueprint` | 分析任務、建立藍圖、分配 Agent | Manager |
| `/claim {agent-id} {task-id}` | 領取任務、自動初始化角色 | Agent |
| `/integrate {task-id}` | 智能整合所有 Agent 的成果 | Super Agent |

### 相關資源位置

| 資源 | 路徑 |
|------|------|
| 藍圖範本 | `.agent-workspace/blueprints/_TEMPLATE.md` |
| 副本工作區 | `.agent-workspace/working/{agent-id}/{task-id}/` |
| Agent 配置檔 | `.claude/agents/{BACKEND,UI-UX,SUPER,MANAGER}.md` |

---

## 使用方式

1. **執行任務前**：讀取此索引，比對觸發條件
2. **找到匹配**：讀取對應的資源檔案
3. **變更 skills/**：完成後更新此索引的檔案清單與更新日期
