# 任務規格：Home 頁面帳號設定介面

> 任務ID：2026-01-08-user-profile-ui
> 指派：UI-UX Agent
> 優先級：High
> 依賴：2026-01-08-user-profile-backend（需先完成 UserProfileHelper.vb）

---

## 目標

1. 在 Home 頁面新增帳號顯示（可點擊開啟設定）
2. 實作「必填欄位彈窗」（登入後首次進入 Home 時檢查）
3. 實作「帳號設定彈窗」（點擊帳號開啟）

---

## 設計參考

### 右上角帳號樣式（參考請購單/費用申請單）

位置：`ExpenseClaimForm.aspx` 第 905-909 行

```html
<div class="site-user-info">
    <asp:Label ID="lblCurrentUser" runat="server" CssClass="user-name" Text=""></asp:Label>
    <span class="separator">｜</span>
    <asp:LinkButton ID="lnkLogout" runat="server" OnClick="lnkLogout_Click">登出</asp:LinkButton>
</div>
```

### 費用部門彈窗樣式（參考費用申請單）

位置：`ExpenseClaimForm.aspx` 第 1720-1744 行

---

## 需求詳情

### 1. Home 頁面帳號顯示

在 `Home.aspx` 新增帳號區域：

```
┌────────────────────────────────────────────┐
│  J E T                    [帳號名] ｜ 登出  │
├────────────────────────────────────────────┤
│                                            │
│              Welcome                       │
│              {UserName}                    │
│                                            │
└────────────────────────────────────────────┘
```

- `[帳號名]` 可點擊，點擊後開啟「帳號設定彈窗」
- 樣式參考 ExpenseClaimForm 的 `.site-user-info`

### 2. 必填欄位彈窗

**觸發條件**：Page_Load 時檢查，若有未填的必填欄位則自動彈出

**彈窗內容**：
- 標題：「請完成帳號設定」
- 說明文字：「以下欄位為必填，請完成設定：」
- 欄位：
  - 費用部門（DropDownList，從 jDEPT 載入）- **僅當未設定時顯示輸入框**
  - 工號（TextBox）- **僅當未設定時顯示輸入框**
- 已填寫的欄位以唯讀方式顯示（Label）
- 確定按鈕

**行為**：
- 若兩個都沒填 → 兩個輸入框都顯示
- 若只缺一個 → 已填的顯示為 Label（不可編輯），缺的顯示輸入框
- 按確定後儲存，若仍有空值則重新顯示彈窗

### 3. 帳號設定彈窗

**觸發條件**：點擊右上角帳號名稱

**彈窗內容**：
- 標題：「帳號設定」
- 欄位（全部可編輯）：
  - 密碼（PasswordBox x2：新密碼、確認密碼）
  - 費用部門（DropDownList）
  - 工號（TextBox）
  - Email（TextBox）
- 儲存按鈕、取消按鈕

**驗證**：
- 密碼：若有輸入，兩個欄位必須一致
- 費用部門、工號：必填
- Email：選填，但若輸入需符合 email 格式

---

## 修改檔案

1. `Home.aspx` - 新增 header、彈窗 HTML
2. `Home.aspx.vb` - 頁面邏輯、調用 UserProfileHelper
3. `Home.aspx.designer.vb` - 新增控制項宣告

---

## 樣式要求

- 使用現有 modalPopup 樣式（參考 ExpenseClaimForm）
- Home 頁面維持深色主題 (`theme-dark`)
- header 樣式可獨立定義，與現有 welcome-panel 協調

---

## 技術提示

1. 使用 AjaxToolkit 的 ModalPopupExtender
2. 調用 `UserProfileHelper.GetMissingRequiredFields()` 檢查必填
3. 調用 `UserProfileHelper.GetExpDeptList()` 載入費用部門選項
4. 調用 `UserProfileHelper.UpdateXxx()` 方法更新資料

---

## 驗收標準

1. [ ] Home 頁面右上角顯示帳號名稱
2. [ ] 點擊帳號可開啟帳號設定彈窗
3. [ ] 登入後若有必填欄位未填，自動彈出必填欄位彈窗
4. [ ] 必填欄位彈窗：已填欄位以 Label 顯示，未填欄位以輸入框顯示
5. [ ] 帳號設定可成功儲存密碼、費用部門、工號、Email
6. [ ] 控制項宣告同步更新至 designer.vb

---

## 完成後

1. 在此目錄建立 `output.md` 說明完成項目
2. 通知 Manager Agent 進行審查
