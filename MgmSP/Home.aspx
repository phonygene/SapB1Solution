<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/MySite1.Master" CodeBehind="Home.aspx.vb" Inherits="MgmSP.Home" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,300;0,400;0,600;0,700;1,300;1,400;1,600;1,700&family=DM+Sans:wght@400;500&display=swap" rel="stylesheet" />

    <style>
        .content-area {
            /*background: transparent;*/
            background: linear-gradient(160deg, #171c30 0%, #1e2440 30%, #252d48 50%, #1e2440 70%, #171c30 100%);
            display: flex;
            justify-content: center;
            align-items: center;
        }
        .welcome-panel {
            text-align: center;
            opacity: 1 !important;
          filter: none !important;
          text-shadow: none !important;
          mix-blend-mode: normal !important;
        }
        .welcome-title {
            font-family: "Cormorant Garamond", Georgia, serif;
            font-size: 26px;
            font-weight: 600;
            font-style: italic;
            letter-spacing: 0.35em;
            /*color: rgba(180, 190, 215, 0.6);*/
            /*color: #1a1f35;*/
            color: #f0f8ff;
            text-transform: uppercase;
            margin-bottom: 16px;
            opacity: 1 !important;
          filter: none !important;
          text-shadow: none !important;
          mix-blend-mode: normal !important;
        }
        .welcome-user {
            font-family: "Cormorant Garamond", Georgia, serif;
            font-size: 36px;
            font-weight: 700;
            /*color: rgba(235, 240, 250, 0.85);*/
            /*color: #1a1f35;*/
            color: #f0f8ff;
            letter-spacing: 0.08em;
            opacity: 1 !important;
          filter: none !important;
          text-shadow: none !important;
          mix-blend-mode: normal !important;
        }

        /* ========================================
           使用者資訊面板 (右上角)
           ======================================== */
        .user-info-panel {
            position: fixed;
            top: 16px;
            right: 24px;
            display: flex;
            align-items: center;
            gap: 12px;
            z-index: 100;
        }

        .user-info-panel .user-name {
            color: rgba(235, 240, 250, 0.9);
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            padding: 8px 16px;
            border-radius: 6px;
            background: rgba(255, 255, 255, 0.08);
            border: 1px solid rgba(255, 255, 255, 0.12);
            transition: all 0.2s ease;
            text-decoration: none;
        }

        .user-info-panel .user-name:hover {
            background: rgba(255, 255, 255, 0.15);
            border-color: rgba(255, 255, 255, 0.25);
            color: #ffffff;
        }

        .user-info-panel .logout-link {
            color: rgba(235, 240, 250, 0.7);
            font-size: 12px;
            text-decoration: none;
            padding: 6px 12px;
            border-radius: 4px;
            transition: all 0.2s ease;
        }

        .user-info-panel .logout-link:hover {
            color: #ffffff;
            background: rgba(255, 255, 255, 0.1);
        }

        /* ========================================
           Modal 樣式
           ======================================== */
        .modalBackground {
            background-color: rgba(0, 0, 0, 0.6);
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            z-index: 999;
        }

        .modalPopup {
            background: #ffffff;
            border-radius: 12px;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
            min-width: 420px;
            max-width: 500px;
            z-index: 1000;
        }

        .modalHeader {
            background: linear-gradient(135deg, #2c3e50 0%, #3498db 100%);
            color: #ffffff;
            padding: 16px 20px;
            border-radius: 12px 12px 0 0;
            display: flex;
            justify-content: space-between;
            align-items: center;
            font-size: 16px;
            font-weight: 600;
        }

        .modalHeader .close-btn {
            color: #ffffff;
            text-decoration: none;
            font-size: 18px;
            font-weight: bold;
            opacity: 0.8;
            transition: opacity 0.2s ease;
        }

        .modalHeader .close-btn:hover {
            opacity: 1;
        }

        .modalBody {
            padding: 24px;
        }

        .modalFooter {
            padding: 16px 24px;
            border-top: 1px solid #e0e0e0;
            display: flex;
            justify-content: flex-end;
            gap: 12px;
            border-radius: 0 0 12px 12px;
            background: #f8f9fa;
        }

        /* 表單樣式 */
        .form-row {
            margin-bottom: 16px;
        }

        .form-row label {
            display: block;
            margin-bottom: 6px;
            font-size: 13px;
            font-weight: 500;
            color: #495057;
        }

        .form-row label .required {
            color: #dc3545;
            margin-left: 2px;
        }

        .form-row input[type="text"],
        .form-row select {
            width: 100%;
            padding: 10px 12px;
            border: 1px solid #ced4da;
            border-radius: 6px;
            font-size: 14px;
            transition: border-color 0.2s ease, box-shadow 0.2s ease;
        }

        .form-row input[type="text"]:focus,
        .form-row select:focus {
            outline: none;
            border-color: #3498db;
            box-shadow: 0 0 0 3px rgba(52, 152, 219, 0.15);
        }

        .form-row input[readonly] {
            background-color: #e9ecef;
            cursor: not-allowed;
        }

        .form-row .field-hint {
            font-size: 11px;
            color: #6c757d;
            margin-top: 4px;
        }

        .form-row .field-error {
            font-size: 11px;
            color: #dc3545;
            margin-top: 4px;
            display: none;
        }

        .form-row.has-error input,
        .form-row.has-error select {
            border-color: #dc3545;
        }

        .form-row.has-error .field-error {
            display: block;
        }

        /* 按鈕樣式 */
        .btn {
            padding: 10px 20px;
            border: none;
            border-radius: 6px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s ease;
        }

        .btn-primary {
            background: linear-gradient(135deg, #3498db 0%, #2980b9 100%);
            color: #ffffff;
        }

        .btn-primary:hover {
            background: linear-gradient(135deg, #2980b9 0%, #1f6dad 100%);
            box-shadow: 0 4px 12px rgba(52, 152, 219, 0.3);
        }

        .btn-secondary {
            background: #e9ecef;
            color: #495057;
        }

        .btn-secondary:hover {
            background: #dee2e6;
        }

        /* 訊息提示 */
        .save-message {
            padding: 10px 16px;
            border-radius: 6px;
            margin-bottom: 16px;
            font-size: 13px;
            display: none;
        }

        .save-message.success {
            background: #d4edda;
            color: #155724;
            border: 1px solid #c3e6cb;
            display: block;
        }

        .save-message.error {
            background: #f8d7da;
            color: #721c24;
            border: 1px solid #f5c6cb;
            display: block;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server"></asp:ScriptManager>

    <!-- 右上角使用者資訊面板 -->
    <div class="user-info-panel">
        <asp:LinkButton ID="lnkUserSettings" runat="server" CssClass="user-name"
            OnClick="lnkUserSettings_Click">
            <asp:Label ID="lblUserDisplay" runat="server" Text=""></asp:Label>
        </asp:LinkButton>
        <asp:HyperLink ID="lnkLogout" runat="server" NavigateUrl="~/usermgm/logout.aspx"
            CssClass="logout-link">登出</asp:HyperLink>
    </div>

    <!-- 歡迎面板 -->
    <div class="welcome-panel">
        <div class="welcome-title">Welcome</div>
        <div class="welcome-user"><asp:Label ID="lblUserName" runat="server" Text=""></asp:Label></div>
    </div>

    <!-- 帳號設定 Modal -->
    <asp:Button ID="btnUserSettingsDummy" runat="server" Style="display:none" />
    <ajaxToolkit:ModalPopupExtender ID="mpeUserSettings" runat="server"
        BehaviorID="mpeUserSettingsBehavior"
        TargetControlID="btnUserSettingsDummy"
        PopupControlID="pnlUserSettings"
        BackgroundCssClass="modalBackground"
        DropShadow="false" />

    <asp:Panel ID="pnlUserSettings" runat="server" CssClass="modalPopup" Style="display:none;">
        <div class="modalHeader">
            <span>帳號設定</span>
            <asp:LinkButton ID="btnCloseSettings" runat="server" CssClass="close-btn"
                OnClick="btnCloseSettings_Click">✕</asp:LinkButton>
        </div>
        <asp:UpdatePanel ID="upUserSettings" runat="server">
            <ContentTemplate>
                <div class="modalBody">
                    <!-- 訊息提示 -->
                    <asp:Panel ID="pnlMessage" runat="server" CssClass="save-message" Visible="false">
                        <asp:Label ID="lblMessage" runat="server"></asp:Label>
                    </asp:Panel>

                    <!-- 帳號 (唯讀) -->
                    <div class="form-row">
                        <label>帳號</label>
                        <asp:TextBox ID="txtUserId" runat="server" ReadOnly="true"></asp:TextBox>
                        <div class="field-hint">帳號無法修改</div>
                    </div>

                    <!-- 姓名 -->
                    <div class="form-row">
                        <label>姓名 <span class="required">*</span></label>
                        <asp:TextBox ID="txtUserName" runat="server" MaxLength="50"></asp:TextBox>
                        <asp:Label ID="lblNameError" runat="server" CssClass="field-error" Visible="false"></asp:Label>
                    </div>

                    <!-- Email -->
                    <div class="form-row">
                        <label>Email</label>
                        <asp:TextBox ID="txtEmail" runat="server" MaxLength="100"></asp:TextBox>
                        <asp:Label ID="lblEmailError" runat="server" CssClass="field-error" Visible="false"></asp:Label>
                    </div>

                    <!-- ＊費用部門 -->
                    <div class="form-row">
                        <label>＊費用部門 <span class="required">*</span></label>
                        <asp:DropDownList ID="ddlExpDept" runat="server">
                        </asp:DropDownList>
                        <asp:Label ID="lblExpDeptError" runat="server" CssClass="field-error" Visible="false"></asp:Label>
                    </div>

                    <!-- ＊工號 -->
                    <div class="form-row">
                        <label>＊工號 <span class="required">*</span></label>
                        <asp:DropDownList ID="ddlEmpSeries" runat="server">
                        </asp:DropDownList>
                        <asp:Label ID="lblEmpSeriesError" runat="server" CssClass="field-error" Visible="false"></asp:Label>
                        <div class="field-hint">SAP 員工系列編號</div>
                    </div>
                </div>
                <div class="modalFooter">
                    <asp:Button ID="btnCancelSettings" runat="server" Text="取消"
                        CssClass="btn btn-secondary" OnClick="btnCancelSettings_Click" />
                    <asp:Button ID="btnSaveSettings" runat="server" Text="儲存"
                        CssClass="btn btn-primary" OnClick="btnSaveSettings_Click" />
                </div>
            </ContentTemplate>
        </asp:UpdatePanel>
    </asp:Panel>
</asp:Content>
