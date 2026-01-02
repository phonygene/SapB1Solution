# 環境健康檢查與故障排除指南

**版本**：1.0
**建立日期**：2025-12-01
**目的**：確保 Agent 協作時的讀寫操作穩定，並提供遭遇問題時的標準排除流程。

---

## 1. 初始環境檢查 (Project Initialization)

Agent 首次進入專案或開始新的 Session 時，應優先檢查以下設定：

### 1.1 檔案編碼與格式標準化
- **檢查目標**：確認專案根目錄是否存在 `.editorconfig`。
- **重要性**：**極高**。這是避免 `apply_diff` 失敗的第一道防線。
- **標準設定**：
  - Windows/ASP.NET 專案應強制使用 `CRLF` 換行。
  - `.vb`, `.cs`, `.aspx` 等傳統檔案應使用 `utf-8-bom` 編碼。
  - 現代 Web 檔案 (`.js`, `.json`) 應使用 `utf-8` 編碼。

若 `.editorconfig` 缺失，請建立以下標準版本：
```ini
root = true

[*]
charset = utf-8-bom
end_of_line = crlf
indent_style = space
indent_size = 4
insert_final_newline = true
trim_trailing_whitespace = true

[*.{xml,config,json,yml,yaml,md,html,css,js,ts}]
charset = utf-8
indent_size = 2

[*.{vb,cs,aspx,ashx,asax}]
indent_size = 4
charset = utf-8-bom
```

### 1.2 Web.config 一致性檢查
- **檢查目標**：`web.config` 中的 `<globalization fileEncoding="..." />` 設定。
- **潛在風險**：若設定為 `BIG5` 但檔案實際為 `UTF-8`，雖不影響 Agent 讀寫，但可能導致 Runtime 亂碼。
- **行動**：記錄此設定，並在讀取原始碼出現亂碼時優先懷疑此處。

---

## 2. 讀寫失敗排除流程 (Troubleshooting)

當 `read_file` 出現亂碼或 `apply_diff` 失敗時，請依序執行：

### 步驟一：確認檔案實際編碼
- 讀取檔案的前幾行，觀察是否有亂碼。
- 若有亂碼，嘗試推測是否為 `Big5` (CP950) 與 `UTF-8` 的誤判。

### 步驟二：確認換行符號 (Line Endings)
- `apply_diff` 對換行符號極度敏感。
- **症狀**：`SEARCH` 區塊看起來完全正確，但 Agent 回報 "Search content not found"。
- **解法**：
    1. 讀取檔案確認當前狀態。
    2. 檢查 `.editorconfig` 是否正確強制 `end_of_line = crlf`。
    3. 若問題持續，嘗試擴大 `SEARCH` 範圍或分段替換。

### 步驟三：驗證寫入權限與鎖定
- 檢查是否有其他程序（如 IIS Express, Visual Studio Debugger）鎖定檔案。
- 嘗試對檔案進行無害修改（如添加註解）以驗證寫入通道是否暢通。

---

## 3. 預防措施 (Best Practices)

1. **始終遵守 .editorconfig**：不要手動覆蓋編輯器的編碼設定。
2. **使用 UTF-8 with BOM 於舊專案**：Visual Studio 對無 BOM 的 UTF-8 支援度不一，保持 BOM 可減少問題。
3. **保持原子性修改**：`apply_diff` 盡量針對單一功能區塊，避免一次跨越多個函式的大範圍修改，以降低上下文匹配失敗的機率。