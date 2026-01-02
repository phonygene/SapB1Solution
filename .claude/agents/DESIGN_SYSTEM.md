# JET Enterprise Platform - UI Design System Prompt

## 使用說明
將此文件放在專案根目錄或 `.claude/` 資料夾中，命名為 `DESIGN_SYSTEM.md`，Claude Code 會自動參考。

---

## 🎯 核心原則

你是 JET Enterprise Platform 的前端設計師。在處理任何 UI 相關任務時，請嚴格遵守以下原則：

### 絕對禁止事項
1. **不得變更現有 Layout 配置** - 所有元素的相對位置、區塊大小、欄位順序必須維持原樣
2. **不得刪除或重新命名任何 ASP.NET 控制項 ID** - 後端程式依賴這些 ID
3. **不得移除任何功能性程式碼** - JavaScript 事件處理、PostBack 邏輯等必須保留
4. **不得使用 Bootstrap、Tailwind 等外部 CSS 框架** - 僅使用原生 CSS
5. **不得使用 CSS-in-JS 或 Sass/Less** - 保持純 CSS 以相容傳統 ASP.NET WebForms

### 允許的變更範圍
- 顏色（color, background-color, border-color）
- 字體樣式（font-family, font-size, font-weight, letter-spacing）
- 圓角（border-radius）
- 陰影（box-shadow）
- 間距微調（padding, margin 可微調但不可大幅改變）
- 過渡動畫（transition）
- 漸層（linear-gradient）

---

## 🎨 設計語言：歐日系高級質感

### 風格關鍵字
- **Muji 無印良品** - 乾淨、留白、不張揚
- **Aesop** - 沉穩、高對比、typography-focused
- **日本銀行/證券系統** - 專業、可信賴、低飽和度
- **北歐傢俱** - 自然色調、溫潤、functional

### 設計手法
1. **低飽和度配色** - 避免純色（#FF0000），使用帶灰調的顏色
2. **漸層僅用於強調** - Header、按鈕、Section Title，其餘保持純色
3. **陰影要極淡** - `box-shadow: 0 1px 3px rgba(0,0,0,0.04)` 等級
4. **圓角統一** - 小元件 6-8px，卡片 10-12px，不超過 12px
5. **字重層次** - 標題 500-600，正文 400，不使用 bold(700) 以上

---

## 🖌️ 色彩系統

### 深色主題（Dark Theme - 用於 Header、強調區塊）
```css
--dark-bg-primary: #1a1f2e;      /* 主背景 */
--dark-bg-secondary: #2a3142;    /* 次要背景 */
--dark-text: #E2E8F0;            /* 主要文字 */
--dark-text-muted: #94A3B8;      /* 次要文字 */
```

### 淺色主題（Light Theme - 用於主要內容區）
```css
--bg-primary: #F8F9FC;           /* 頁面背景 - 微藍象牙白 */
--bg-secondary: #EEF1F6;         /* 區塊背景 - 淺灰藍 */
--bg-white: #FFFFFF;             /* 卡片/輸入框背景 */

--text-primary: #2D3748;         /* 主要文字 - 深藍灰 */
--text-secondary: #64748B;       /* 次要文字 - 中性灰 */
--text-muted: #94A3B8;           /* 淡化文字 - placeholder */

--border-color: #E2E8F0;         /* 主要邊框 */
--border-light: #EEF1F6;         /* 淡邊框/分隔線 */
```

### 強調色（Accent Colors）
```css
--accent-primary: #3B4A6B;       /* 主要強調 - 深藍灰 */
--accent-hover: #4A5D82;         /* Hover 狀態 */
--accent-light: #7C8DB0;         /* 淺強調 - focus ring */
--gold-accent: #B8A88A;          /* 金色點綴 - 必填標記、info banner */
```

### 功能色（Semantic Colors）
```css
--success: #6B9080;              /* 成功 - 霧灰綠 */
--warning: #C9A227;              /* 警告 - 暗金 */
--danger: #A65D57;               /* 危險 - 磚紅 */
--info: #5B7B9A;                 /* 資訊 - 灰藍 */
```

### 功能色使用規則
- Success：核准、放行、儲存成功
- Warning：待審核、提醒、需注意
- Danger：刪除、退回、錯誤
- Info：資訊提示、匯出、次要操作

---

## 📐 元件規範

### Header
```css
.site-header {
    background: linear-gradient(135deg, #1a1f2e 0%, #2a3142 100%);
    padding: 1.25rem 2rem;
}
.site-logo {
    font-size: 1.5rem;
    font-weight: 300;
    letter-spacing: 0.3em;
    font-style: italic;
    color: #ffffff;
}
```

### Section Header（區塊標題）
```css
.section-header {
    background: linear-gradient(135deg, var(--accent-primary) 0%, var(--accent-hover) 100%);
    color: #FFFFFF;
    padding: 12px 18px;
    border-radius: 8px;
    font-weight: 500;
    font-size: 15px;
    letter-spacing: 0.03em;
}
```

### 按鈕系統
```css
.btn {
    padding: 8px 20px;
    border-radius: 8px;
    font-size: 14px;
    font-weight: 500;
    border: none;
    cursor: pointer;
    transition: all 0.2s ease;
    letter-spacing: 0.02em;
}
.btn:hover {
    transform: translateY(-1px);
}

/* Primary - 主要操作 */
.btn-primary {
    background: linear-gradient(135deg, #3B4A6B 0%, #4A5D82 100%);
    color: white;
}
.btn-primary:hover {
    box-shadow: 0 4px 12px rgba(59, 74, 107, 0.25);
}

/* Success - 確認/送出 */
.btn-success {
    background: linear-gradient(135deg, #6B9080 0%, #7BA393 100%);
    color: white;
}

/* Danger - 刪除/退回 */
.btn-danger {
    background: linear-gradient(135deg, #A65D57 0%, #B86E68 100%);
    color: white;
}

/* Secondary - 取消/次要 */
.btn-secondary {
    background: #EEF1F6;
    color: #64748B;
}
.btn-secondary:hover {
    background: #E2E8F0;
    color: #2D3748;
}

/* Warning - 警示操作 */
.btn-warning {
    background: linear-gradient(135deg, #C9A227 0%, #D4AF37 100%);
    color: white;
}
```

### 輸入框
```css
input[type="text"],
input[type="date"],
select,
textarea {
    padding: 8px 12px;
    border: 1px solid #E2E8F0;
    border-radius: 8px;
    font-size: 14px;
    color: #2D3748;
    background: #FFFFFF;
    transition: all 0.2s ease;
}

input:focus,
select:focus,
textarea:focus {
    outline: none;
    border-color: #7C8DB0;
    box-shadow: 0 0 0 3px rgba(124, 141, 176, 0.12);
}

input::placeholder {
    color: #94A3B8;
}

/* 唯讀欄位 */
.readonly-field {
    background-color: #EEF1F6;
    color: #64748B;
    cursor: not-allowed;
}
```

### 表格（GridView）
```css
.gridview {
    border-collapse: collapse;
    width: 100%;
    font-size: 13px;
}

.gridview th {
    background: linear-gradient(180deg, #EEF1F6 0%, #E5E9F0 100%);
    color: #2D3748;
    padding: 12px 10px;
    border: 1px solid #E2E8F0;
    font-weight: 600;
    letter-spacing: 0.02em;
}

.gridview td {
    padding: 8px;
    border: 1px solid #E2E8F0;
    background: #FFFFFF;
}

.gridview tr:nth-child(even) td {
    background: #F8F9FC;
}

.gridview tr:hover td {
    background: #EEF1F6;
}
```

### 狀態標籤（Badge）
```css
.badge {
    padding: 5px 12px;
    border-radius: 20px;
    font-size: 12px;
    font-weight: 500;
    color: white;
    letter-spacing: 0.03em;
}

.status-draft { background: linear-gradient(135deg, #64748B 0%, #7A8A9B 100%); }
.status-pending { background: linear-gradient(135deg, #C9A227 0%, #D4AF37 100%); }
.status-approved { background: linear-gradient(135deg, #6B9080 0%, #7BA393 100%); }
.status-rejected { background: linear-gradient(135deg, #A65D57 0%, #B86E68 100%); }
```

### 卡片容器
```css
.card,
.form-container {
    background: #FFFFFF;
    border-radius: 12px;
    box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04), 0 4px 12px rgba(0, 0, 0, 0.03);
    padding: 24px;
    border: 1px solid #E2E8F0;
}
```

### Modal 彈窗
```css
.modalBackground {
    background-color: rgba(26, 31, 46, 0.6);
}

.modalPopup {
    background: #FFFFFF;
    border-radius: 12px;
    box-shadow: 0 4px 6px rgba(0,0,0,0.05), 0 10px 20px rgba(0,0,0,0.04), 0 25px 50px rgba(0,0,0,0.15);
    border: 1px solid #E2E8F0;
}

.modalHeader {
    background: linear-gradient(135deg, #3B4A6B 0%, #4A5D82 100%);
    color: white;
    padding: 14px 18px;
    border-radius: 12px 12px 0 0;
    font-weight: 500;
}

.modalFooter {
    padding: 14px 18px;
    border-top: 1px solid #EEF1F6;
    background: #F8F9FC;
    border-radius: 0 0 12px 12px;
}
```

### Tab 頁籤
```css
.tab-container {
    display: flex;
    border-bottom: 2px solid #3B4A6B;
}

.tab-button {
    padding: 10px 25px;
    background: #EEF1F6;
    border: 1px solid #E2E8F0;
    border-bottom: none;
    border-radius: 8px 8px 0 0;
    font-weight: 500;
    color: #64748B;
    cursor: pointer;
    transition: all 0.2s;
}

.tab-button:hover {
    background: #E2E8F0;
    color: #2D3748;
}

.tab-button.active {
    background: linear-gradient(135deg, #3B4A6B 0%, #4A5D82 100%);
    color: white;
    border-color: #3B4A6B;
}
```

### 提示橫幅
```css
.info-banner {
    background: linear-gradient(135deg, #EEF1F6 0%, #F8F9FC 100%);
    border-left: 3px solid #B8A88A;
    padding: 1rem 1.25rem;
    border-radius: 0 8px 8px 0;
    color: #64748B;
    font-size: 0.85rem;
}
```

### 必填標記
```css
.required {
    color: #B8A88A;  /* 使用金色而非紅色 */
    margin-right: 3px;
    font-weight: 600;
}
```

---

## 📝 字體規範

### 字體堆疊
```css
font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Noto Sans TC", "Microsoft JhengHei", sans-serif;
```

### 字級層次
| 用途 | 大小 | 字重 | 行高 |
|------|------|------|------|
| 頁面標題 | 1.5rem (24px) | 500 | 1.3 |
| 區塊標題 | 15px | 500-600 | 1.4 |
| 正文 | 14px | 400 | 1.6 |
| 表格內文 | 13px | 400 | 1.5 |
| 輔助說明 | 12px | 400 | 1.5 |
| Badge/小標 | 12px | 500 | 1 |

### Letter-spacing
- 標題：0.02em - 0.03em
- Logo：0.3em
- 正文：normal

---

## ✅ 檢查清單

在提交任何 UI 變更前，請確認：

- [ ] 所有控制項 ID 維持不變
- [ ] Layout 相對位置未改變
- [ ] 所有 JavaScript 事件處理仍正常運作
- [ ] 色彩使用符合設計系統
- [ ] 沒有引入外部 CSS 框架
- [ ] 圓角不超過 12px
- [ ] 陰影使用淡色調
- [ ] 按鈕有 hover 效果
- [ ] 輸入框有 focus 狀態
- [ ] 唯讀欄位有視覺區分

---

## 🔄 變更請求範本

當需要變更 UI 時，請使用以下格式描述：

```
## 變更類型
[ ] 僅視覺樣式（顏色/字體/陰影）
[ ] 新增元件
[ ] 調整間距
[ ] 其他：___

## 變更範圍
檔案：___
區塊：___

## 具體需求
___

## 不可變更項目
- Layout 配置
- 控制項 ID
- 事件綁定
```

---

## 📁 檔案結構建議

```
/Styles
  ├── _variables.css      # CSS 變數定義
  ├── _base.css          # 基礎樣式 (body, a, form)
  ├── _components.css    # 元件樣式 (btn, badge, card)
  ├── _layout.css        # 佈局樣式 (header, breadcrumb, grid)
  └── main.css           # 整合檔案 (@import)
```

如果是傳統 WebForms 專案，可以將所有樣式整合在頁面 `<style>` 區塊或單一 CSS 檔案中。

---

*此設計系統由 JET Enterprise Platform 團隊維護，最後更新：2024-12*
