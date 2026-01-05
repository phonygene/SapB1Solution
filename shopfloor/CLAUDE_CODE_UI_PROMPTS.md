# Claude Code UI 設計任務 Prompt 範本

## 使用方式
複製下方 Prompt 並根據實際需求修改，然後貼到 Claude Code 中執行。

---

## 🎨 通用視覺重設計 Prompt

```
請為 [檔案路徑] 套用 JET Enterprise Platform 的視覺設計系統。

## 設計風格
歐日系高級質感，參考 Muji、Aesop 的設計語言：
- 低飽和度配色，避免純色
- 極淡陰影 (rgba 透明度 0.03-0.05)
- 統一圓角 8px（卡片 12px）
- 漸層僅用於 Header、按鈕、區塊標題

## 色彩系統
深色區塊（Header）：#1a1f2e → #2a3142 漸層
頁面背景：#F8F9FC（微藍象牙白）
卡片背景：#FFFFFF
主要文字：#2D3748
次要文字：#64748B
主強調色：#3B4A6B（深藍灰）
邊框色：#E2E8F0
金色點綴：#B8A88A（用於必填標記、info banner）

功能色：
- Success: #6B9080
- Warning: #C9A227  
- Danger: #A65D57
- Info: #5B7B9A

## 嚴格限制
1. ❌ 不得變更任何 Layout 配置（元素位置、區塊大小、欄位順序）
2. ❌ 不得變更任何控制項 ID 或 Name
3. ❌ 不得刪除或修改任何 JavaScript 程式碼
4. ❌ 不得引入 Bootstrap、Tailwind 等外部框架
5. ✅ 僅變更 CSS 樣式（顏色、字體、圓角、陰影、過渡效果）

## 輸出要求
- 保留原始檔案的完整結構
- 只修改 <style> 區塊內的 CSS
- 在修改處加上註解標記變更原因
```

---

## 🔧 單一元件樣式調整 Prompt

```
請調整 [檔案路徑] 中的 [元件名稱] 樣式。

## 當前問題
[描述目前的視覺問題]

## 期望效果
[描述期望的視覺效果]

## 設計規範
- 主色：#3B4A6B
- 背景：#F8F9FC
- 邊框：#E2E8F0
- 圓角：8px
- 陰影：0 1px 3px rgba(0,0,0,0.04)

## 限制
- 僅修改指定元件的 CSS
- 不變更 HTML 結構
- 不影響其他元件樣式
```

---

## 🆕 新增元件 Prompt

```
請在 [檔案路徑] 中新增 [元件描述]。

## 元件需求
[詳細描述元件功能與內容]

## 插入位置
在 [參考元素] 的 [之前/之後]

## 設計規範
遵循 JET Design System：
- 背景：#FFFFFF（卡片）或 #F8F9FC（頁面）
- 邊框：1px solid #E2E8F0
- 圓角：12px（卡片）、8px（按鈕/輸入框）
- 標題：linear-gradient(135deg, #3B4A6B, #4A5D82)
- 按鈕：使用 .btn-primary / .btn-secondary / .btn-success / .btn-danger

## 必要條件
- 為新元件加上適當的 ASP.NET 控制項 ID（遵循現有命名慣例）
- 確保 RWD 相容性
- 新增必要的 CSS class
```

---

## 🎯 完整頁面重設計 Prompt（進階版）

```
請為 [檔案路徑] 進行完整視覺重設計。

## 專案背景
這是 JET Enterprise Platform 的 [頁面功能描述]，使用 ASP.NET WebForms + AjaxControlToolkit。

## 設計目標
打造歐日系高級企業軟體質感，風格參考：
- Muji 無印良品的乾淨留白
- Aesop 的沉穩對比
- 日本金融系統的專業可信賴感

## 完整色彩配置

### CSS 變數定義（請加在 style 最前面）
:root {
    --bg-primary: #F8F9FC;
    --bg-secondary: #EEF1F6;
    --bg-white: #FFFFFF;
    --text-primary: #2D3748;
    --text-secondary: #64748B;
    --text-muted: #94A3B8;
    --accent-primary: #3B4A6B;
    --accent-hover: #4A5D82;
    --accent-light: #7C8DB0;
    --border-color: #E2E8F0;
    --border-light: #EEF1F6;
    --gold-accent: #B8A88A;
    --success: #6B9080;
    --warning: #C9A227;
    --danger: #A65D57;
    --info: #5B7B9A;
    --shadow-sm: 0 1px 3px rgba(0,0,0,0.04), 0 4px 12px rgba(0,0,0,0.03);
    --shadow-md: 0 4px 6px rgba(0,0,0,0.05), 0 10px 20px rgba(0,0,0,0.04);
}

## 元件樣式規範

### 按鈕
- Primary: 漸層 #3B4A6B → #4A5D82，白字
- Success: 漸層 #6B9080 → #7BA393，白字
- Danger: 漸層 #A65D57 → #B86E68，白字
- Secondary: 背景 #EEF1F6，文字 #64748B
- Hover: transform: translateY(-1px) + box-shadow
- 圓角: 8px，padding: 8px 20px

### 輸入框
- 邊框: 1px solid #E2E8F0
- Focus: border-color #7C8DB0 + box-shadow 0 0 0 3px rgba(124,141,176,0.12)
- 圓角: 8px
- Placeholder: #94A3B8

### 表格
- 表頭: 漸層背景 #EEF1F6 → #E5E9F0
- 斑馬紋: nth-child(even) #F8F9FC
- Hover: #EEF1F6
- 邊框: 1px solid #E2E8F0

### Modal
- 背景遮罩: rgba(26, 31, 46, 0.6)
- 彈窗圓角: 12px
- Header: 漸層 #3B4A6B → #4A5D82
- Footer: 背景 #F8F9FC

## 嚴格執行事項
1. 所有現有控制項 ID 必須保留不變
2. 所有現有 JavaScript 函數必須保留不變  
3. Layout 結構（row, col-half, form-group）位置不變
4. GridView 的欄位順序不變
5. 僅修改視覺呈現相關的 CSS 屬性

## 變更記錄
完成後請列出所有變更的 CSS class 和屬性。
```

---

## 💡 常用追加指令

### 微調色彩
```
請將 [元件] 的主色從 #3B4A6B 調整為 [新色碼]，並同步更新相關的 hover、focus 狀態。
```

### 增加動畫效果
```
請為 [元件] 增加細微的過渡動畫，使用 transition: all 0.2s ease，hover 時增加 translateY(-1px) 效果。
```

### 統一圓角
```
請檢查所有元件的 border-radius，統一為：按鈕/輸入框 8px，卡片 12px，badge 20px。
```

### 調整陰影
```
請將所有陰影調整為更淡的效果：box-shadow: 0 1px 3px rgba(0,0,0,0.04), 0 4px 12px rgba(0,0,0,0.03)
```

---

## ⚠️ 故障排除

如果 Claude Code 仍然亂改東西，嘗試加上這段：

```
## 重要提醒
在進行任何修改前，請先：
1. 完整閱讀現有程式碼結構
2. 識別所有 ASP.NET 控制項 (asp:Button, asp:TextBox 等)
3. 識別所有 JavaScript 函數和事件綁定
4. 確認 UpdatePanel 的範圍

修改時請：
- 使用 str_replace 而非重寫整個檔案
- 每次只修改一小段 CSS
- 修改後確認沒有破壞原有功能

如果不確定某段程式碼的作用，請先詢問而非直接修改。
```
