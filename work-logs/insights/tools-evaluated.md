# 工具評估紀錄

> 記錄評估過的工具、函式庫、技術方案

---

## [2026-01-02] CSS 主題系統方案

### 評估選項
1. CSS Class 切換 (`.theme-blue`, `.theme-green`)
2. CSS Custom Properties + data-theme 屬性
3. CSS-in-JS 方案

### 結論：採用選項 2

### 原因
- data-theme 屬性語義清晰
- 不需要額外 JavaScript 框架
- 與現有 ASP.NET Web Forms 相容
- 支援 12 種主題，易於擴展

### 實作方式
```html
<body data-theme="green">
```
```css
[data-theme="green"] {
  --accent-primary: #3D5A4C;
  ...
}
```

---

## [2026-01-02] 多 Agent 協作方案

### 評估選項
1. AI 開會討論
2. 自由工作 + 最後 Merge
3. Manager 協調 + 結構化審查
4. 混合模式

### 結論：採用選項 4 (混合模式)

### 原因
- AI 開會目前效率不高，容易冗長且缺乏真正碰撞
- 純自由工作可能有大量 Merge 衝突
- 完全由 Manager 協調可能成為瓶頸
- 混合模式：
  - 獨立任務自由執行
  - 有相依性時 Manager 協調
  - 完成後結構化審查

### 實作方式
- task-status.json 追蹤狀態和衝突
- work-logs 記錄經驗
- 定期反思優化
