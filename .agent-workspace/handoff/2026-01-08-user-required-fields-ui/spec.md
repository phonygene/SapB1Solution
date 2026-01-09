# UI-UX 任務規格：使用者必填資訊 - 介面設計

> 任務 ID: 2026-01-08-user-required-fields-ui
> 指派: UI-UX Agent
> 優先級: High
> 建立時間: 2026-01-08
> 相依: 2026-01-08-user-required-fields-backend（Phase 1）

---

## 目標

在 Home.aspx 建立使用者帳號資訊區塊和設定介面，並修改費用部門彈窗 UI。

---

## 任務清單

### U1: Home.aspx 右上角使用者資訊區塊

位置：內容區域右上角（參考費用申請單的單據編號區塊風格）

```html
<div class="user-info-panel">
    <div class="user-name">
        <asp:LinkButton ID="lnkUserName" runat="server"
            OnClick="lnkUserName_Click">使用者名稱</asp:LinkButton>
    </div>
    <div class="user-id">帳號: xxx</div>
</div>
```

樣式建議：
- 位置：右上角，不遮擋 Welcome 訊息
- 風格：與現有主題一致，使用 CSS 變數
- 使用者名稱：可點擊超連結樣式

### U2: 使用者設定 Modal

使用 AjaxControlToolkit ModalPopupExtender（與費用部門彈窗一致）

```html
<asp:Panel ID="pnlUserProfile" runat="server" CssClass="modalPopup" Style="display:none; width:450px;">
    <div class="modalHeader">
        <span>帳號設定</span>
    </div>
    <div class="modalBody">
        <!-- 密碼 -->
        <div class="form-group">
            <label class="form-label">新密碼:</label>
            <div class="form-control">
                <asp:TextBox ID="txtNewPassword" runat="server"
                    TextMode="Password" placeholder="留空則不變更"></asp:TextBox>
            </div>
        </div>
        <div class="form-group">
            <label class="form-label">確認密碼:</label>
            <div class="form-control">
                <asp:TextBox ID="txtConfirmPassword" runat="server"
                    TextMode="Password"></asp:TextBox>
            </div>
        </div>

        <!-- 工號 (必填) -->
        <div class="form-group">
            <label class="form-label"><span class="required">*</span>工號:</label>
            <div class="form-control">
                <asp:TextBox ID="txtEmpSeries" runat="server"></asp:TextBox>
            </div>
        </div>

        <!-- 費用部門 (必填) -->
        <div class="form-group">
            <label class="form-label"><span class="required">*</span>費用部門:</label>
            <div class="form-control">
                <asp:DropDownList ID="ddlExpDept" runat="server"></asp:DropDownList>
            </div>
        </div>

        <!-- Email -->
        <div class="form-group">
            <label class="form-label">Email:</label>
            <div class="form-control">
                <asp:TextBox ID="txtEmail" runat="server"
                    TextMode="Email"></asp:TextBox>
            </div>
        </div>
    </div>
    <div class="modalFooter">
        <asp:Button ID="btnSaveProfile" runat="server" Text="儲存"
            CssClass="btn btn-primary" OnClick="btnSaveProfile_Click" />
        <asp:Button ID="btnCancelProfile" runat="server" Text="取消"
            CssClass="btn btn-secondary" />
    </div>
</asp:Panel>
```

### U3: 前端驗證

- 密碼：若有輸入，兩次必須一致
- 工號：必填
- 費用部門：必填
- Email：格式驗證（可選填）

### U4: Home.aspx.vb 後端整合

```vb
' 頁面載入時
Protected Sub Page_Load(...) Handles Me.Load
    If Not IsPostBack Then
        LoadUserProfile()
    End If
End Sub

Private Sub LoadUserProfile()
    Dim profile = UserProfileHelper.GetUserProfile(currentUserId)
    lnkUserName.Text = profile.UserName
    lblUserId.Text = profile.UserId
    ' ...
End Sub

' 點擊使用者名稱
Protected Sub lnkUserName_Click(...)
    LoadExpDeptDropDown()
    PopulateProfileForm()
    mpeUserProfile.Show()
End Sub

' 儲存
Protected Sub btnSaveProfile_Click(...)
    ' 驗證
    ' 呼叫 UserProfileHelper.UpdateUserProfile()
    ' 關閉彈窗
End Sub
```

### U5: 修改費用部門彈窗 UI（ExpenseClaimForm.aspx）

在現有的 `pnlExpDept` Panel 中新增工號輸入欄位：

```html
<asp:Panel ID="pnlExpDept" runat="server" CssClass="modalPopup" Style="display:none; width:400px;">
    <div class="modalHeader" style="background: linear-gradient(135deg, #5B7B9A 0%, #6B8BA9 100%);">
        <span>設定必填資訊</span>  <!-- 標題改為更通用 -->
    </div>
    <div class="modalBody">
        <p style="margin-bottom:15px; color: var(--text-secondary);">
            請完成以下必填資訊設定：
        </p>

        <!-- 新增：工號欄位 -->
        <div class="form-group">
            <label class="form-label" style="width:100px;"><span class="required">*</span>工號:</label>
            <div class="form-control">
                <asp:TextBox ID="txtEmpSeriesPopup" runat="server"
                    placeholder="請輸入您的工號"></asp:TextBox>
            </div>
        </div>

        <!-- 現有：費用部門欄位 -->
        <div class="form-group">
            <label class="form-label" style="width:100px;"><span class="required">*</span>費用部門:</label>
            <div class="form-control">
                <asp:DropDownList ID="ddlExpDeptSelect" runat="server" Width="100%">
                </asp:DropDownList>
            </div>
        </div>
    </div>
    <div class="modalFooter">
        <asp:Button ID="btnExpDeptConfirm" runat="server" Text="確定"
            CssClass="btn btn-primary" OnClick="btnExpDeptConfirm_Click" />
    </div>
</asp:Panel>
```

---

## 設計規範

1. **樣式一致性**：使用現有 CSS 變數（--accent-primary, --text-secondary 等）
2. **RWD**：Modal 在小螢幕時自適應寬度
3. **必填標記**：紅色星號 `<span class="required">*</span>`
4. **密碼顯示**：使用 `TextMode="Password"` 顯示黑色圓點

---

## 影響檔案

- `MgmSP/Home.aspx`
- `MgmSP/Home.aspx.vb`
- `MgmSP/Home.aspx.designer.vb`（新控制項宣告）
- `MgmSP/ExpenseClaimForm.aspx`（彈窗 UI）
- `MgmSP/ExpenseClaimForm.aspx.designer.vb`（新控制項宣告）

---

## 相依性

- **等待 Backend 完成**：UserProfileHelper.vb 模組
- **Phase 2 可並行**：Home.aspx 的 UI 可先行設計，後端呼叫待 Helper 完成後整合

---

## 驗收標準

1. [ ] Home.aspx 右上角顯示使用者名稱（可點擊）和帳號
2. [ ] 點擊後開啟設定 Modal，正確顯示現有資料
3. [ ] 密碼欄位顯示為黑色圓點
4. [ ] 儲存後資料正確更新
5. [ ] 費用部門彈窗包含工號輸入欄位
6. [ ] 所有新控制項已在 designer.vb 中宣告

---

## 完成後

1. 更新 `.agent-workspace/handoff/2026-01-08-user-required-fields-ui/output.md`
2. 通知 Manager 進行審查
