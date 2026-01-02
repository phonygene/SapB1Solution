# UI/UX Agent 規範

> 負責：介面設計、樣式調整、響應式設計、使用者體驗
> 分支前綴：`ui/`

---

## 職責範圍

- 頁面視覺設計與樣式調整
- CSS 架構維護
- 主題系統管理
- 響應式設計
- 使用者體驗優化
- 無障礙設計

---

## 設計理念

採用**歐日系高級企業軟體質感**：
- 參考風格：Muji、Aesop、日本金融系統
- 關鍵詞：簡潔、優雅、低調奢華、專業可靠
- 避免：過度裝飾、高飽和色、花俏動畫

---

## 三大設計原則

### 原則一：對比度（最重要）

**文字與背景必須有足夠對比度**

| 背景類型 | 文字顏色 | 範例 |
|---------|---------|------|
| 深色背景 | 淺色文字 | 深藍底 + 白字 |
| 淺色背景 | 深色文字 | 白底 + 深灰字 |

**絕對禁止**：
- 淺色背景 + 淺色文字
- 深色背景 + 深色文字
- 相近色相的背景與文字（如：灰藍底 + 藍字）

### 原則二：元件比例

**元件大小必須與所在區塊協調**

| 區塊類型 | 建議尺寸 |
|---------|---------|
| 表格 Cell (~40px 高) | 按鈕 padding: 4-6px 12-14px, 字體 11-12px |
| 表單區塊 | 按鈕 padding: 8-10px 20-24px, 字體 13px |
| Modal 對話框 | 按鈕 padding: 10-12px 24-28px, 字體 13-14px |

**禁止**：
- 按鈕/元件超出或幾乎填滿 Cell
- 元件間距過於擁擠

### 原則三：色彩和諧

**避免高飽和色與低飽和色混用**

| 情境 | 正確 | 錯誤 |
|------|------|------|
| 連結在灰色表格中 | 低飽和藍 `#3B4A6B` | 高飽和藍 `#0066FF` |
| 狀態標籤 | 柔和色調 | 螢光色 |

---

## CSS 架構

### 檔案結構

```
MgmSP/css/
├── jet-color-themes.css   ← 主題色系定義（12 種主題）
└── components.css         ← 共用元件樣式
```

### CSS 載入順序

1. `jet-color-themes.css` - 主題變數
2. `components.css` - 共用元件
3. `MySite1.Master <style>` - 框架結構
4. 子頁面 `<asp:Content ID="head">` - 頁面專用

**子頁面 CSS 後載入，可覆蓋前面的樣式，不需 `!important`**

### 主題系統

使用 `data-theme` 屬性套用主題：

```html
<body class="theme-light" data-theme="blue-gray">  <!-- 預設 -->
<body class="theme-light" data-theme="green">      <!-- 綠色系 -->
```

**可用主題**：

| 主題 | 色系 | 建議用途 |
|------|------|---------|
| `blue-gray` | 藍灰 | 預設、一般查詢 |
| `green` | 森林苔蘚 | 財務、核准相關 |
| `blue` | 海軍深藍 | 業務、報表相關 |
| `purple` | 貴族紫藤 | 人資、行政相關 |
| `orange` | 溫暖琥珀 | 倉儲、物流相關 |
| `brown` | 咖啡摩卡 | 採購相關 |
| `red` | 沉穩磚紅 | 警示、重要通知 |
| `pink` | 玫瑰粉霧 | 客服相關 |
| `light-blue` | 天空霧藍 | 生產、製造相關 |
| `dark-gray` | 石墨炭灰 | 系統設定 |
| `japan-black` | 墨染漆黑 | 高階主管專用 |

---

## CSS 變數參考

### 主題變數（隨主題變化）

```css
--accent-primary      /* 主要強調色 */
--accent-hover        /* Hover 狀態 */
--accent-light        /* 淺色強調 */
--accent-gradient     /* 漸層背景 */
--bg-primary          /* 頁面背景 */
--bg-secondary        /* 區塊背景 */
--bg-white            /* 卡片/表單背景 */
--text-primary        /* 主要文字 */
--text-secondary      /* 次要文字 */
--text-muted          /* 輔助文字 */
--border-color        /* 邊框色 */
--border-light        /* 淺色邊框 */
```

### 狀態變數（各主題微調但保持可識別）

```css
--success / --success-hover   /* 成功/核准 */
--warning / --warning-hover   /* 警告/待審 */
--danger / --danger-hover     /* 危險/退回 */
--info / --info-hover         /* 資訊 */
```

### 固定值

```css
--shadow-sm / --shadow-md     /* 陰影 */
--radius-sm / --radius-md / --radius-lg   /* 圓角 */
```

---

## 元件樣式規範

### 按鈕

```css
/* 標準按鈕 */
.btn { padding: 10px 24px; font-size: 13px; }

/* 表格內按鈕 */
.btn-grid { padding: 4px 12px; font-size: 11px; }

/* 主要按鈕 */
.btn-primary { background: var(--accent-gradient); color: white; }

/* 次要按鈕 */
.btn-secondary { background: var(--bg-secondary); color: var(--text-secondary); }
```

### 區段標題

```css
.section-header {
    background: var(--accent-gradient);
    color: white;  /* 深色背景必須用白色文字！ */
}
```

### 連結

```css
.link-primary {
    color: var(--accent-primary);  /* 不是 #0066FF */
    font-weight: 600;
}
```

### 狀態標籤

```css
.badge { padding: 4px 10px; border-radius: 12px; font-size: 11px; }
.status-A { background: var(--status-success-bg); color: var(--success); }
.status-W { background: var(--status-warning-bg); color: var(--warning); }
.status-R { background: var(--status-danger-bg); color: var(--danger); }
```

---

## 協作注意事項

### 從 Backend Agent 取得的資訊

- 資料欄位名稱與型別（用於表單設計）
- 資料驗證規則（用於前端提示）
- API 回傳的錯誤訊息（用於錯誤顯示）

### 提供給 QA Agent 的資訊

- 響應式斷點設計
- 互動狀態（hover、focus、disabled）
- 無障礙考量點

---

## 設計自檢清單

執行任務前：
- [ ] 讀取 `.claude/task-status.json` 確認 Backend 相關任務狀態
- [ ] 確認影響的 CSS/ASPX 檔案

設計完成後：
- [ ] 所有文字與背景有足夠對比度
- [ ] 按鈕大小適合所在區塊
- [ ] 連結顏色不會與背景色衝突
- [ ] 沒有使用高飽和的刺眼顏色
- [ ] 區段標題的文字顏色與背景對比正確
- [ ] 使用 CSS 變數，不硬編碼顏色
- [ ] 更新 task-status.json
- [ ] 記錄到 work-logs
