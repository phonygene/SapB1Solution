<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="PurchaseRequestForm.aspx.vb" Inherits="MgmSP.PurchaseRequestForm"
    MaintainScrollPositionOnPostback="true" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">

<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>請購單</title>
    <style type="text/css">
        /* ========================================
           JET Enterprise Platform - Elegant Theme
           ======================================== */

        :root {
            /* 主色系 */
            --bg-primary: #F8F9FC;
            --bg-secondary: #EEF1F6;
            --bg-white: #FFFFFF;

            /* 文字色 */
            --text-primary: #2D3748;
            --text-secondary: #64748B;
            --text-muted: #94A3B8;

            /* 強調色 - 綠色系 */
            --accent-primary: #4A6B5B;
            --accent-hover: #5B7B6B;
            --accent-light: #7B9B8B;

            /* 功能色 */
            --border-color: #E2E8F0;
            --border-light: #EEF1F6;
            --gold-accent: #B8A88A;
            --success: #6B9080;
            --warning: #C9A227;
            --danger: #A65D57;
            --info: #5B7B9A;

            /* 陰影 */
            --shadow-sm: 0 1px 3px rgba(0, 0, 0, 0.04), 0 4px 12px rgba(0, 0, 0, 0.03);
            --shadow-md: 0 4px 6px rgba(0, 0, 0, 0.05), 0 10px 20px rgba(0, 0, 0, 0.04);
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Noto Sans TC", "Microsoft JhengHei", sans-serif;
            font-size: 14px;
            background-color: var(--bg-primary);
            color: var(--text-primary);
            line-height: 1.6;
        }

        .form-container {
            max-width: 1400px;
            margin: 20px auto;
            padding: 24px;
            background-color: var(--bg-white);
            box-shadow: var(--shadow-sm);
            border-radius: 12px;
            border: 1px solid var(--border-color);
        }

        .section-header {
            background: linear-gradient(135deg, var(--accent-primary) 0%, var(--accent-hover) 100%);
            color: #FFFFFF;
            padding: 12px 18px;
            margin: 24px 0 18px 0;
            font-weight: 500;
            border-radius: 8px;
            font-size: 15px;
            letter-spacing: 0.03em;
        }

        /* Layout Grid */
        .row {
            display: flex;
            flex-wrap: wrap;
            margin-right: -15px;
            margin-left: -15px;
        }

        .col-half {
            flex: 0 0 50%;
            max-width: 50%;
            padding-right: 15px;
            padding-left: 15px;
            box-sizing: border-box;
        }

        .form-group {
            margin-bottom: 14px;
            display: flex;
            align-items: center;
        }

        .form-label {
            width: 140px;
            font-weight: 500;
            padding-right: 10px;
            text-align: right;
            color: var(--text-primary);
        }

        .form-control {
            flex: 1;
            display: flex;
            align-items: center;
        }

        input[type="text"],
        input[type="date"],
        input[type="number"],
        select,
        textarea {
            padding: 8px 12px;
            border: 1px solid var(--border-color);
            border-radius: 8px;
            font-family: inherit;
            font-size: 14px;
            width: 100%;
            box-sizing: border-box;
            color: var(--text-primary);
            background-color: var(--bg-white);
            transition: all 0.2s ease;
        }

        input[type="text"]:focus,
        input[type="date"]:focus,
        input[type="number"]:focus,
        select:focus,
        textarea:focus {
            outline: none;
            border-color: var(--accent-light);
            box-shadow: 0 0 0 3px rgba(123, 155, 139, 0.15);
        }

        textarea {
            resize: vertical;
        }

        .readonly-field {
            background-color: var(--bg-secondary);
            color: var(--text-secondary);
            cursor: not-allowed;
            border-color: var(--border-light);
        }

        /* 按鈕系統 */
        .btn {
            padding: 8px 20px;
            border-radius: 8px;
            border: none;
            cursor: pointer;
            font-size: 14px;
            font-weight: 500;
            margin-right: 5px;
            transition: all 0.2s ease;
            letter-spacing: 0.02em;
        }

        .btn-primary {
            background: linear-gradient(135deg, var(--accent-primary) 0%, var(--accent-hover) 100%);
            color: white;
        }

        .btn-primary:hover {
            background: linear-gradient(135deg, var(--accent-hover) 0%, #6B8B7B 100%);
            box-shadow: 0 4px 12px rgba(74, 107, 91, 0.25);
        }

        .btn-success {
            background: linear-gradient(135deg, var(--success) 0%, #7BA393 100%);
            color: white;
        }

        .btn-success:hover {
            box-shadow: 0 4px 12px rgba(107, 144, 128, 0.3);
        }

        .btn-danger {
            background: linear-gradient(135deg, var(--danger) 0%, #B86E68 100%);
            color: white;
        }

        .btn-danger:hover {
            box-shadow: 0 4px 12px rgba(166, 93, 87, 0.3);
        }

        .btn-secondary {
            background: var(--bg-secondary);
            color: var(--text-secondary);
        }

        .btn-secondary:hover {
            background: var(--border-color);
            color: var(--text-primary);
        }

        .btn-warning {
            background: linear-gradient(135deg, var(--warning) 0%, #D4AF37 100%);
            color: white;
        }

        .btn-info {
            background: linear-gradient(135deg, var(--info) 0%, #6B8BA9 100%);
            color: white;
        }

        .btn:hover {
            opacity: 1;
            transform: translateY(-1px);
        }

        .btn-icon {
            padding: 6px 12px;
            font-size: 12px;
            margin-left: 5px;
        }

        .required {
            color: var(--gold-accent);
            margin-right: 3px;
            font-weight: 600;
        }

        .error-text {
            color: var(--danger);
            font-size: 12px;
            margin-left: 5px;
            display: block;
        }

        /* GridView */
        .gridview {
            border-collapse: collapse;
            width: 100%;
            margin-top: 10px;
            font-size: 13px;
        }

        .gridview th {
            background: linear-gradient(180deg, var(--bg-secondary) 0%, #E5E9F0 100%);
            color: var(--text-primary);
            padding: 12px 10px;
            border: 1px solid var(--border-color);
            text-align: center;
            white-space: nowrap;
            font-weight: 600;
            letter-spacing: 0.02em;
        }

        .gridview td {
            padding: 8px;
            border: 1px solid var(--border-color);
            vertical-align: middle;
            background-color: var(--bg-white);
        }

        .gridview tr:nth-child(even) td {
            background-color: var(--bg-primary);
        }

        .gridview tr:hover td {
            background-color: var(--bg-secondary);
        }

        .gridview input[type="text"],
        .gridview input[type="number"],
        .gridview select {
            width: 95%;
            padding: 6px 8px;
        }

        /* Status Badges */
        .badge {
            padding: 5px 12px;
            border-radius: 20px;
            color: white;
            font-size: 12px;
            font-weight: 500;
            letter-spacing: 0.03em;
        }

        .status-P { background: linear-gradient(135deg, var(--text-secondary) 0%, #7A8A9B 100%); }
        .status-W { background: linear-gradient(135deg, var(--warning) 0%, #D4AF37 100%); color: white; }
        .status-A { background: linear-gradient(135deg, var(--success) 0%, #7BA393 100%); }
        .status-R { background: linear-gradient(135deg, var(--danger) 0%, #B86E68 100%); }

        /* Modal */
        .modalBackground {
            background-color: rgba(26, 31, 46, 0.6);
        }

        .modalPopup {
            background-color: var(--bg-white);
            border-radius: 12px;
            padding: 0;
            width: 700px;
            box-shadow: var(--shadow-md), 0 25px 50px rgba(0, 0, 0, 0.15);
            border: 1px solid var(--border-color);
        }

        .modalHeader {
            background: linear-gradient(135deg, var(--accent-primary) 0%, var(--accent-hover) 100%);
            color: white;
            padding: 14px 18px;
            border-radius: 12px 12px 0 0;
            font-weight: 500;
            display: flex;
            justify-content: space-between;
            align-items: center;
            letter-spacing: 0.03em;
        }

        .modalBody {
            padding: 20px;
            max-height: 500px;
            overflow-y: auto;
        }

        .modalFooter {
            padding: 14px 18px;
            border-top: 1px solid var(--border-light);
            text-align: right;
            background: var(--bg-primary);
            border-radius: 0 0 12px 12px;
        }

        /* Header & Breadcrumb */
        .site-header {
            background: linear-gradient(135deg, #1a1f2e 0%, #2a3142 100%);
            padding: 1rem 2rem;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }

        .site-logo {
            font-size: 1.5rem;
            font-weight: 300;
            color: #ffffff;
            letter-spacing: 0.3em;
            font-style: italic;
            text-decoration: none;
        }

        .site-logo:hover {
            color: #ffffff;
        }

        .site-user-info {
            color: var(--text-muted);
            font-size: 0.875rem;
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }

        .site-user-info .user-name {
            color: #E2E8F0;
            font-weight: 500;
        }

        .site-user-info .separator {
            color: var(--text-muted);
        }

        .site-user-info a {
            color: var(--text-muted);
            text-decoration: none;
            transition: color 0.2s ease;
        }

        .site-user-info a:hover {
            color: #E2E8F0;
        }

        .breadcrumb {
            padding: 0.875rem 2rem;
            background: var(--bg-white);
            border-bottom: 1px solid var(--border-color);
            font-size: 0.875rem;
            color: var(--text-secondary);
        }

        .breadcrumb a {
            color: var(--accent-primary);
            text-decoration: none;
            transition: color 0.2s ease;
        }

        .breadcrumb a:hover {
            color: var(--accent-hover);
        }

        .breadcrumb .separator {
            margin: 0 0.5rem;
            color: var(--text-muted);
        }

        .breadcrumb .current {
            color: var(--text-secondary);
        }

        /* 頁面標題區 */
        .page-header-title {
            margin: 0;
            color: var(--accent-primary);
            font-weight: 500;
            letter-spacing: 0.02em;
        }

        /* 審核區塊 */
        .approval-panel {
            margin-top: 24px;
            padding: 20px;
            background: linear-gradient(135deg, #FFFBEB 0%, #FEF8E8 100%);
            border: 1px solid var(--gold-accent);
            border-radius: 10px;
            border-left: 4px solid var(--gold-accent);
        }

        .approval-panel h3 {
            margin-top: 0;
            color: var(--accent-primary);
            font-weight: 500;
        }

        /* Footer 分隔線 */
        .footer-section {
            margin-top: 24px;
            border-top: 1px solid var(--border-color);
            padding-top: 16px;
        }

        /* 總額顯示 */
        .total-amount {
            font-weight: 600;
            font-size: 22px;
            color: var(--accent-primary);
        }

        .total-detail {
            color: var(--text-muted);
            font-size: 12px;
        }

        /* 搜尋按鈕組合 */
        .search-combo {
            display: flex;
            width: 100%;
        }

        .search-combo input {
            border-top-right-radius: 0;
            border-bottom-right-radius: 0;
        }

        .search-combo .btn {
            border-top-left-radius: 0;
            border-bottom-left-radius: 0;
            margin: 0;
        }

        /* 供應商資訊提示 */
        .vendor-info {
            margin-top: 6px;
            font-size: 12px;
            color: var(--accent-light);
        }

        /* Empty data 提示 */
        .empty-data-hint {
            text-align: center;
            padding: 24px;
            color: var(--text-muted);
        }

        /* Radio Button List */
        input[type="radio"] {
            accent-color: var(--accent-primary);
        }

        /* Checkbox */
        input[type="checkbox"] {
            accent-color: var(--accent-primary);
            width: 16px;
            height: 16px;
        }

        /* 訊息標籤 */
        .message-label {
            font-weight: 500;
        }

        /* Pager */
        .gridview a {
            padding: 6px 12px;
            margin: 0 2px;
            border: 1px solid var(--border-color);
            border-radius: 6px;
            color: var(--accent-primary);
            background: var(--bg-white);
            display: inline-block;
        }

        .gridview a:hover {
            background: var(--bg-secondary);
        }

        .gridview span {
            padding: 6px 12px;
            margin: 0 2px;
            background: var(--accent-primary);
            color: white;
            border-radius: 6px;
            display: inline-block;
        }

        /* 驗證結果彈窗 */
        .validation-modal {
            width: 500px;
        }

        .validation-section {
            margin-bottom: 15px;
        }

        .validation-section-title {
            font-weight: 500;
            padding: 10px 14px;
            border-radius: 6px;
            margin-bottom: 10px;
        }

        .validation-section-title.error {
            background: linear-gradient(135deg, #FDF2F2 0%, #FEE8E8 100%);
            color: var(--danger);
            border-left: 4px solid var(--danger);
        }

        .validation-section-title.warning {
            background: linear-gradient(135deg, #FFFBEB 0%, #FEF3CD 100%);
            color: #92700C;
            border-left: 4px solid var(--warning);
        }

        .validation-list {
            list-style: none;
            padding: 0;
            margin: 0 0 0 16px;
        }

        .validation-list li {
            padding: 6px 0;
            border-bottom: 1px dashed var(--border-light);
            color: var(--text-secondary);
        }

        .validation-list li:last-child {
            border-bottom: none;
        }

        .validation-list li:before {
            content: "• ";
            font-weight: bold;
        }

        .validation-list.error li:before {
            color: var(--danger);
        }

        .validation-list.warning li:before {
            color: var(--warning);
        }

        /* 附件上傳區 */
        .file-upload-area {
            display: flex;
            align-items: center;
        }

        /* 連結樣式 */
        a {
            color: var(--accent-primary);
            text-decoration: none;
            transition: color 0.2s ease;
        }

        a:hover {
            color: var(--accent-hover);
        }
    </style>
    <script type="text/javascript">
        function confirmDelete() {
            return confirm('確定要刪除此筆請購單嗎？此操作無法復原。');
        }

        // 防止按鈕連續點擊
        var isSubmitting = false;

        function preventDoubleClick(btn, msg) {
            if (typeof (Sys) !== 'undefined' && Sys.WebForms && Sys.WebForms.PageRequestManager) {
                var prm = Sys.WebForms.PageRequestManager.getInstance();
                if (!prm.get_isInAsyncPostBack() && isSubmitting) {
                    console.log('Resetting stuck isSubmitting flag');
                    isSubmitting = false;
                }
            }

            if (isSubmitting) {
                return false;
            }

            if (typeof (Page_ClientValidate) == 'function') {
                if (Page_ClientValidate() == false) {
                    return false;
                }
            }

            if (msg && !confirm(msg)) {
                return false;
            }

            isSubmitting = true;
            setTimeout(function () {
                if (isSubmitting) {
                    btn.disabled = true;
                    btn.value = '處理中...';
                }
            }, 50);
            return true;
        }

        function initPageLogic() {
            if (typeof (Sys) !== 'undefined' && Sys.WebForms && Sys.WebForms.PageRequestManager) {
                var prm = Sys.WebForms.PageRequestManager.getInstance();

                prm.add_endRequest(function () {
                    isSubmitting = false;
                    var btns = document.querySelectorAll('.btn');
                    for (var i = 0; i < btns.length; i++) {
                        btns[i].disabled = false;
                    }
                });
            }
        }

        if (window.addEventListener) {
            window.addEventListener('load', initPageLogic);
        } else if (window.attachEvent) {
            window.attachEvent('onload', initPageLogic);
        }

        function toggleGridCheckboxes(source, gridId, checkboxIdSuffix) {
            var grid = document.getElementById(gridId);
            if (!grid) {
                return;
            }
            var inputs = grid.getElementsByTagName('input');
            for (var i = 0; i < inputs.length; i++) {
                var input = inputs[i];
                if (input.type === 'checkbox' && input.id && input.id.indexOf(checkboxIdSuffix) >= 0) {
                    if (input !== source) {
                        input.checked = source.checked;
                    }
                }
            }
        }
    </script>
</head>

<body>
    <form id="form1" runat="server">
        <!-- Site Header -->
        <header class="site-header">
            <a href="Home.aspx?smid=index&smode=0" class="site-logo">J E T</a>
            <div class="site-user-info">
                <asp:Label ID="lblCurrentUser" runat="server" CssClass="user-name" Text=""></asp:Label>
                <span class="separator">｜</span>
                <asp:LinkButton ID="lnkLogout" runat="server" OnClick="lnkLogout_Click">登出</asp:LinkButton>
            </div>
        </header>

        <!-- Breadcrumb Navigation -->
        <nav class="breadcrumb">
            <asp:HyperLink ID="lnkHome" runat="server" NavigateUrl="~/Home.aspx?smid=index&smode=0">首頁</asp:HyperLink>
            <span class="separator">></span>
            <span class="current">請購單</span>
        </nav>

        <asp:ScriptManager ID="ScriptManager1" runat="server"></asp:ScriptManager>
        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
            <ContentTemplate>
                <asp:HiddenField ID="hfRateDate" runat="server" Value="" />
                <div class="form-container">
                    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:20px;">
                        <h2 class="page-header-title">請購單 (Purchase Request)</h2>
                        <div style="text-align:right;">
                            <asp:Label ID="lblDocNum" runat="server" Text="[New]" Font-Bold="True"
                                Font-Size="18px" ForeColor="#4A6B5B"></asp:Label>
                            <br />
                            <asp:Label ID="lblDocStatus" runat="server" CssClass="badge status-P" Text="草稿">
                            </asp:Label>
                        </div>
                    </div>

                    <!-- Header Section -->
                    <div class="section-header">表頭資訊</div>

                    <div class="row">
                        <!-- Left Column -->
                        <div class="col-half">
                            <div class="form-group">
                                <label class="form-label"><span class="required">*</span>請購人:</label>
                                <div class="form-control">
                                    <asp:TextBox ID="txtReqName" runat="server" ReadOnly="true" CssClass="readonly-field"></asp:TextBox>
                                </div>
                            </div>
                            <div class="form-group">
                                <label class="form-label">請購部門:</label>
                                <div class="form-control">
                                    <asp:DropDownList ID="ddlReqDept" runat="server"></asp:DropDownList>
                                </div>
                            </div>
                            <div class="form-group">
                                <label class="form-label">建議供應商代碼:</label>
                                <div class="form-control">
                                    <div class="search-combo">
                                        <asp:TextBox ID="txtCardCode" runat="server" placeholder="請點選搜尋"></asp:TextBox>
                                        <asp:Button ID="btnSearchCardCode" runat="server" Text="🔍"
                                            CssClass="btn btn-secondary"
                                            OnClick="btnSearchCardCode_Click" />
                                    </div>
                                    <div class="vendor-info">
                                        <asp:Label ID="lblVendorInfo" runat="server"></asp:Label>
                                    </div>
                                </div>
                            </div>
                            <div class="form-group">
                                <label class="form-label">建議供應商名稱:</label>
                                <div class="form-control">
                                    <div class="search-combo">
                                        <asp:TextBox ID="txtCardName" runat="server" placeholder="請點選搜尋"></asp:TextBox>
                                        <asp:Button ID="btnSearchCardName" runat="server" Text="🔍"
                                            CssClass="btn btn-secondary"
                                            OnClick="btnSearchCardName_Click" />
                                    </div>
                                </div>
                            </div>
                            <div class="form-group">
                                <label class="form-label"><span class="required">*</span>幣別:</label>
                                <div class="form-control">
                                    <asp:DropDownList ID="ddlDocCurrency" runat="server" AutoPostBack="True"
                                        OnSelectedIndexChanged="ddlDocCurrency_SelectedIndexChanged"
                                        Width="40%" style="margin-right:5px;"></asp:DropDownList>
                                    <asp:TextBox ID="txtDocRate" runat="server" Width="30%" Text="1.0"
                                        AutoPostBack="true" OnTextChanged="txtDocRate_TextChanged"></asp:TextBox>
                                    <asp:Button ID="btnRefreshRate" runat="server" Text="↻"
                                        CssClass="btn btn-secondary btn-icon" OnClick="btnRefreshRate_Click"
                                        ToolTip="更新匯率" />
                                </div>
                            </div>
                        </div>

                        <!-- Right Column -->
                        <div class="col-half">
                            <div class="form-group">
                                <label class="form-label">請購單號 (jID):</label>
                                <div class="form-control">
                                    <asp:TextBox ID="txtJID" runat="server" ReadOnly="true"
                                        CssClass="readonly-field" placeholder="系統自動產生"></asp:TextBox>
                                </div>
                            </div>
                            <div class="form-group">
                                <label class="form-label"><span class="required">*</span>請購日期:</label>
                                <div class="form-control">
                                    <asp:TextBox ID="txtDocDate" runat="server" TextMode="Date"
                                        AutoPostBack="true" OnTextChanged="btnRefreshRate_Click"></asp:TextBox>
                                    <asp:Label ID="lblErrDocDate" runat="server" CssClass="error-text"
                                        Visible="False"></asp:Label>
                                </div>
                            </div>
                            <div class="form-group">
                                <label class="form-label">需求日期:</label>
                                <div class="form-control">
                                    <asp:TextBox ID="txtReqDate" runat="server" TextMode="Date"></asp:TextBox>
                                </div>
                            </div>
                            <div class="form-group">
                                <label class="form-label">文件狀態:</label>
                                <div class="form-control">
                                    <asp:TextBox ID="txtStatusDisplay" runat="server" ReadOnly="true"
                                        CssClass="readonly-field"></asp:TextBox>
                                </div>
                            </div>
                            <div class="form-group">
                                <label class="form-label">放行狀態:</label>
                                <div class="form-control">
                                    <asp:TextBox ID="txtApprovalStatus" runat="server" ReadOnly="true"
                                        CssClass="readonly-field"></asp:TextBox>
                                </div>
                            </div>
                            <div class="form-group">
                                <label class="form-label">簽核系統 PID:</label>
                                <div class="form-control">
                                    <asp:TextBox ID="txtUPID" runat="server"></asp:TextBox>
                                </div>
                            </div>
                        </div>
                    </div>

                    <!-- Detail Section Header -->
                    <div class="section-header">請購明細</div>

                    <div style="margin-bottom: 10px; display:flex; justify-content:space-between;">
                        <div>
                            <asp:Button ID="btnAddLine" runat="server" Text="+ 新增明細"
                                OnClick="btnAddLine_Click" CssClass="btn btn-primary" />
                            <asp:Button ID="btnDeleteLine" runat="server" Text="🗑 刪除選取"
                                OnClick="btnDeleteLine_Click" CssClass="btn btn-danger"
                                OnClientClick="return confirm('確定刪除選中的明細行？');" />
                        </div>
                        <div class="file-upload-area">
                            <asp:FileUpload ID="fileUpload" runat="server"
                                style="display:inline-block; width:200px;" AllowMultiple="true" />
                            <asp:Button ID="btnUpload" runat="server" Text="上傳附件"
                                OnClick="btnUpload_Click" CssClass="btn btn-secondary btn-icon" />
                        </div>
                    </div>

                    <div style="margin-bottom:10px;">
                        <asp:GridView ID="gvAttachments" runat="server" AutoGenerateColumns="False"
                            CssClass="gridview" OnRowCommand="gvAttachments_RowCommand">
                            <Columns>
                                <asp:BoundField DataField="FileName" HeaderText="檔案名稱" />
                                <asp:BoundField DataField="UploadDate" HeaderText="上傳日期"
                                    DataFormatString="{0:yyyy-MM-dd}" />
                                <asp:BoundField DataField="UploadTime" HeaderText="上傳時間" />
                                <asp:BoundField DataField="Uploader" HeaderText="上傳者" />
                                <asp:TemplateField HeaderText="動作">
                                    <ItemTemplate>
                                        <asp:LinkButton ID="lbtnDelete" runat="server"
                                            CommandName="DeleteFile"
                                            CommandArgument='<%# Container.DataItemIndex %>' Text="刪除"
                                            ForeColor="#A65D57" OnClientClick="return confirm('確定刪除此附件？');">
                                        </asp:LinkButton>
                                        <asp:HyperLink ID="hlDownload" runat="server"
                                            NavigateUrl='<%# "DownloadHandler.ashx?id=" & Eval("ID") %>'
                                            Target="_blank" Text="下載" style="margin-left:5px;">
                                        </asp:HyperLink>
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Center" Width="100px" />
                                </asp:TemplateField>
                            </Columns>
                            <EmptyDataTemplate>
                                <div class="empty-data-hint">無附件</div>
                            </EmptyDataTemplate>
                        </asp:GridView>
                    </div>

                    <div style="overflow-x:auto;">
                        <asp:GridView ID="gvPRDetail" runat="server" AutoGenerateColumns="False"
                            CssClass="gridview" OnRowDataBound="gvPRDetail_RowDataBound"
                            OnRowCommand="gvPRDetail_RowCommand">
                            <Columns>
                                <asp:TemplateField HeaderText="選">
                                    <HeaderTemplate>
                                        <input type="checkbox"
                                            onclick="toggleGridCheckboxes(this, '<%= gvPRDetail.ClientID %>', 'chkSelect')" />
                                    </HeaderTemplate>
                                    <ItemTemplate>
                                        <asp:CheckBox ID="chkSelect" runat="server" />
                                    </ItemTemplate>
                                    <ItemStyle Width="30px" HorizontalAlign="Center" />
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="#">
                                    <ItemTemplate>
                                        <asp:Label ID="lblLineNum" runat="server"
                                            Text='<%# Container.DataItemIndex + 1 %>'></asp:Label>
                                    </ItemTemplate>
                                    <ItemStyle Width="40px" HorizontalAlign="Center" />
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="品號">
                                    <ItemTemplate>
                                        <div style="display:flex; align-items:center; gap:4px;">
                                            <asp:TextBox ID="txtItemCode" runat="server" Width="100px"></asp:TextBox>
                                            <asp:Button ID="btnSearchItem" runat="server" Text="🔍"
                                                CssClass="btn btn-secondary btn-icon"
                                                CommandName="SearchItem"
                                                CommandArgument='<%# Container.DataItemIndex %>' />
                                        </div>
                                    </ItemTemplate>
                                    <ItemStyle Width="140px" />
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="品名/說明">
                                    <ItemTemplate>
                                        <asp:TextBox ID="txtDescription" runat="server" Width="180px"></asp:TextBox>
                                    </ItemTemplate>
                                    <ItemStyle Width="190px" />
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="數量">
                                    <ItemTemplate>
                                        <asp:TextBox ID="txtQuantity" runat="server" Width="70px"
                                            style="text-align:right;" AutoPostBack="true"
                                            OnTextChanged="CalculateLineTotal" Text="1"></asp:TextBox>
                                    </ItemTemplate>
                                    <ItemStyle Width="80px" />
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="單價">
                                    <ItemTemplate>
                                        <asp:TextBox ID="txtPrice" runat="server" Width="80px"
                                            style="text-align:right;" AutoPostBack="true"
                                            OnTextChanged="CalculateLineTotal" Text="0"></asp:TextBox>
                                    </ItemTemplate>
                                    <ItemStyle Width="90px" />
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="稅碼">
                                    <ItemTemplate>
                                        <asp:DropDownList ID="ddlVatGroup" runat="server"
                                            Width="80px" AutoPostBack="true"
                                            OnSelectedIndexChanged="CalculateLineTotal">
                                        </asp:DropDownList>
                                    </ItemTemplate>
                                    <ItemStyle Width="90px" />
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="稅額">
                                    <ItemTemplate>
                                        <asp:TextBox ID="txtVatSum" runat="server" Text="0"
                                            Width="70px" style="text-align:right;" ReadOnly="true"
                                            CssClass="readonly-field"></asp:TextBox>
                                    </ItemTemplate>
                                    <ItemStyle Width="80px" />
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="含稅金額">
                                    <ItemTemplate>
                                        <asp:TextBox ID="txtGTotal" runat="server"
                                            Width="90px" style="text-align:right;" ReadOnly="true"
                                            CssClass="readonly-field"></asp:TextBox>
                                    </ItemTemplate>
                                    <ItemStyle Width="100px" />
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="倉庫">
                                    <ItemTemplate>
                                        <asp:DropDownList ID="ddlWhsCode" runat="server" Width="80px">
                                        </asp:DropDownList>
                                    </ItemTemplate>
                                    <ItemStyle Width="90px" />
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="交期">
                                    <ItemTemplate>
                                        <asp:TextBox ID="txtShipDate" runat="server" TextMode="Date" Width="120px"></asp:TextBox>
                                    </ItemTemplate>
                                    <ItemStyle Width="130px" />
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="產品">
                                    <ItemTemplate>
                                        <asp:DropDownList ID="ddlCostingCode" runat="server" Width="80px">
                                        </asp:DropDownList>
                                    </ItemTemplate>
                                    <ItemStyle Width="90px" />
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="部門">
                                    <ItemTemplate>
                                        <asp:DropDownList ID="ddlCostingCode2" runat="server" Width="80px">
                                        </asp:DropDownList>
                                    </ItemTemplate>
                                    <ItemStyle Width="90px" />
                                </asp:TemplateField>
                            </Columns>
                            <EmptyDataTemplate>
                                <div class="empty-data-hint">請新增請購明細</div>
                            </EmptyDataTemplate>
                        </asp:GridView>
                    </div>

                    <!-- Footer Section -->
                    <div class="footer-section">
                        <div class="row">
                            <div class="col-half">
                                <div class="form-group">
                                    <label class="form-label">採購人員:</label>
                                    <div class="form-control">
                                        <asp:DropDownList ID="ddlPurchaser" runat="server"></asp:DropDownList>
                                    </div>
                                </div>
                                <div class="form-group">
                                    <label class="form-label">建立者:</label>
                                    <div class="form-control">
                                        <asp:TextBox ID="txtOwner" runat="server" CssClass="readonly-field"
                                            ReadOnly="true"></asp:TextBox>
                                    </div>
                                </div>
                                <div class="form-group">
                                    <label class="form-label">備註:</label>
                                    <div class="form-control">
                                        <asp:TextBox ID="txtRemarks" runat="server" TextMode="MultiLine"
                                            Height="50px"></asp:TextBox>
                                    </div>
                                </div>
                            </div>
                            <div class="col-half" style="text-align:right;">
                                <div class="form-group" style="justify-content: flex-end;">
                                    <label class="form-label" style="width:auto;">單據總額 (含稅):</label>
                                    <div style="width: 150px; margin-left:10px;">
                                        <asp:Label ID="lblDocTotalWithTax" runat="server" Text="0.00"
                                            CssClass="total-amount"></asp:Label>
                                    </div>
                                </div>
                                <div class="total-detail">
                                    未稅: <asp:Label ID="lblDocTotal" runat="server" Text="0.00"></asp:Label>
                                    |
                                    稅額: <asp:Label ID="lblVatSum" runat="server" Text="0.00"></asp:Label>
                                </div>
                            </div>
                        </div>
                    </div>

                    <!-- Approval Section -->
                    <asp:Panel ID="pnlApproval" runat="server" CssClass="approval-panel">
                        <h3>審核作業</h3>
                        <div class="form-group">
                            <label class="form-label" style="width:100px;">審核意見:</label>
                            <div class="form-control">
                                <asp:TextBox ID="txtApprovalComments" runat="server" TextMode="MultiLine"
                                    Height="60px"></asp:TextBox>
                            </div>
                        </div>
                        <div style="text-align:center; margin-top:10px;">
                            <asp:Button ID="btnApprove" runat="server" Text="放行 (Approve)"
                                OnClick="btnApprove_Click" CssClass="btn btn-success"
                                OnClientClick="return confirm('確定要放行此單據嗎？');" />
                            <asp:Button ID="btnReject" runat="server" Text="退回 (Reject)"
                                OnClick="btnReject_Click" CssClass="btn btn-danger"
                                OnClientClick="return confirm('確定要退回此單據嗎？');" />
                        </div>
                    </asp:Panel>

                    <!-- Buttons -->
                    <div style="text-align:center; margin-top:30px;">
                        <asp:Button ID="btnSubmit" runat="server" Text="儲存並送審 (Save & Submit)"
                            OnClick="btnSubmit_Click" CssClass="btn btn-success"
                            OnClientClick="return preventDoubleClick(this, '確定要送出審核嗎？');" />
                        <asp:Button ID="btnDelete" runat="server" Text="刪除 (Delete)"
                            OnClick="btnDelete_Click" CssClass="btn btn-danger"
                            OnClientClick="return preventDoubleClick(this, null) && confirmDelete();" />
                        <asp:Button ID="btnCancel" runat="server" Text="取消 (Cancel)"
                            OnClick="btnCancel_Click" CssClass="btn btn-secondary" />
                        <asp:Button ID="btnUpdate" runat="server" Text="更新 (Update)"
                            OnClick="btnUpdate_Click" CssClass="btn btn-primary" Visible="false" />
                        <asp:Button ID="btnExportPDF" runat="server" Text="匯出 PDF"
                            OnClick="btnExportPDF_Click" CssClass="btn btn-info" Visible="false" />
                        <asp:Button ID="btnNewDocument" runat="server" Text="新增新單據"
                            OnClick="btnNewDocument_Click" CssClass="btn btn-primary" Visible="false" />

                        <div style="margin-top:10px;">
                            <asp:Label ID="lblMessage" runat="server" CssClass="message-label"></asp:Label>
                        </div>
                    </div>
                </div>

                <!-- Vendor Search Modal -->
                <asp:Button ID="btnDummy" runat="server" style="display:none" />
                <ajaxToolkit:ModalPopupExtender ID="mpeVendor" runat="server" TargetControlID="btnDummy"
                    PopupControlID="pnlVendorSearch" BackgroundCssClass="modalBackground"
                    CancelControlID="btnCloseVendor" />
                <asp:Panel ID="pnlVendorSearch" runat="server" CssClass="modalPopup" style="display:none;">
                    <div class="modalHeader">
                        <span>供應商搜尋</span>
                        <asp:LinkButton ID="btnCloseVendor" runat="server" ForeColor="White"
                            Font-Bold="true" style="text-decoration:none;">✕</asp:LinkButton>
                    </div>
                    <div class="modalBody">
                        <div style="margin-bottom:10px;">
                            <div style="display:flex; align-items:center;">
                                <asp:TextBox ID="txtVendorSearchKeyword" runat="server"
                                    placeholder="輸入關鍵字..."></asp:TextBox>
                                <asp:Button ID="btnDoSearchVendor" runat="server" Text="搜尋"
                                    OnClick="btnDoSearchVendor_Click" CssClass="btn btn-primary"
                                    style="margin-left:5px;" />
                                <asp:HiddenField ID="hfSearchSource" runat="server" />
                            </div>
                            <div style="margin-top:8px;">
                                <asp:RadioButtonList ID="rblSearchMode" runat="server"
                                    RepeatDirection="Horizontal">
                                    <asp:ListItem Value="Fuzzy" Selected="True">模糊搜尋</asp:ListItem>
                                    <asp:ListItem Value="Exact">開頭比對</asp:ListItem>
                                </asp:RadioButtonList>
                            </div>
                        </div>
                        <asp:GridView ID="gvVendorSearch" runat="server" AutoGenerateColumns="False"
                            Width="100%" CssClass="gridview" OnRowCommand="gvVendorSearch_RowCommand"
                            AllowPaging="True" PageSize="10"
                            OnPageIndexChanging="gvVendorSearch_PageIndexChanging">
                            <Columns>
                                <asp:TemplateField HeaderText="動作">
                                    <ItemTemplate>
                                        <asp:LinkButton ID="lbtnSelect" runat="server"
                                            CommandName="SelectVendor"
                                            CommandArgument='<%# Eval("CardCode") + "|" + Eval("CardName") %>'
                                            CssClass="btn btn-success btn-icon">選取</asp:LinkButton>
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Center" Width="70px" />
                                </asp:TemplateField>
                                <asp:BoundField DataField="CardCode" HeaderText="代碼" />
                                <asp:BoundField DataField="CardName" HeaderText="名稱" />
                            </Columns>
                            <PagerStyle HorizontalAlign="Center" CssClass="gridview" />
                        </asp:GridView>
                    </div>
                </asp:Panel>

                <!-- Item Search Modal -->
                <asp:Button ID="btnItemDummy" runat="server" style="display:none" />
                <ajaxToolkit:ModalPopupExtender ID="mpeItem" runat="server"
                    TargetControlID="btnItemDummy"
                    PopupControlID="pnlItemSearch" BackgroundCssClass="modalBackground"
                    CancelControlID="btnCloseItem" />
                <asp:Panel ID="pnlItemSearch" runat="server" CssClass="modalPopup" style="display:none;">
                    <div class="modalHeader">
                        <span>品號搜尋</span>
                        <asp:LinkButton ID="btnCloseItem" runat="server" ForeColor="White"
                            Font-Bold="true" style="text-decoration:none;">✕</asp:LinkButton>
                    </div>
                    <div class="modalBody">
                        <div style="margin-bottom:10px;">
                            <div style="display:flex; align-items:center;">
                                <asp:TextBox ID="txtItemSearchKeyword" runat="server"
                                    placeholder="輸入品號或品名..."></asp:TextBox>
                                <asp:Button ID="btnDoSearchItem" runat="server" Text="搜尋"
                                    OnClick="btnDoSearchItem_Click" CssClass="btn btn-primary"
                                    style="margin-left:5px;" />
                                <asp:HiddenField ID="hfItemSearchRowIndex" runat="server" />
                            </div>
                            <div style="margin-top:8px;">
                                <asp:RadioButtonList ID="rblItemSearchMode" runat="server"
                                    RepeatDirection="Horizontal">
                                    <asp:ListItem Value="Fuzzy" Selected="True">模糊搜尋</asp:ListItem>
                                    <asp:ListItem Value="Exact">開頭比對</asp:ListItem>
                                </asp:RadioButtonList>
                            </div>
                        </div>
                        <asp:GridView ID="gvItemSearch" runat="server" AutoGenerateColumns="False"
                            Width="100%" CssClass="gridview" OnRowCommand="gvItemSearch_RowCommand"
                            AllowPaging="True" PageSize="10"
                            OnPageIndexChanging="gvItemSearch_PageIndexChanging">
                            <Columns>
                                <asp:TemplateField HeaderText="動作">
                                    <ItemTemplate>
                                        <asp:LinkButton ID="lbtnSelectItem" runat="server"
                                            CommandName="SelectItem"
                                            CommandArgument='<%# Eval("ItemCode") + "|" + Eval("ItemName") + "|" + Eval("LastPurPrc") %>'
                                            CssClass="btn btn-success btn-icon">選取</asp:LinkButton>
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Center" Width="70px" />
                                </asp:TemplateField>
                                <asp:BoundField DataField="ItemCode" HeaderText="品號" />
                                <asp:BoundField DataField="ItemName" HeaderText="品名" />
                                <asp:BoundField DataField="LastPurPrc" HeaderText="最近採購價" DataFormatString="{0:N2}" />
                            </Columns>
                            <PagerStyle HorizontalAlign="Center" CssClass="gridview" />
                        </asp:GridView>
                    </div>
                </asp:Panel>

                <!-- 驗證結果彈窗 -->
                <asp:Button ID="btnValidationDummy" runat="server" Style="display:none" />
                <ajaxToolkit:ModalPopupExtender ID="mpeValidation" runat="server"
                    BehaviorID="mpeValidationBehavior" TargetControlID="btnValidationDummy"
                    PopupControlID="pnlValidation" BackgroundCssClass="modalBackground"
                    DropShadow="false" />
                <asp:Panel ID="pnlValidation" runat="server" CssClass="modalPopup validation-modal"
                    Style="display:none;">
                    <div class="modalHeader" style="background: linear-gradient(135deg, #A65D57 0%, #B86E68 100%);" id="divValidationHeader"
                        runat="server">
                        <span>單據檢核結果</span>
                        <asp:LinkButton ID="btnValidationClose" runat="server" ForeColor="White"
                            Font-Bold="true" Style="text-decoration:none;"
                            OnClick="btnValidationBack_Click">✕</asp:LinkButton>
                    </div>
                    <div class="modalBody">
                        <asp:Panel ID="pnlErrors" runat="server" CssClass="validation-section"
                            Visible="false">
                            <div class="validation-section-title error">錯誤 (必須修正才能儲存)</div>
                            <asp:BulletedList ID="blErrors" runat="server" CssClass="validation-list error">
                            </asp:BulletedList>
                        </asp:Panel>
                        <asp:Panel ID="pnlWarnings" runat="server" CssClass="validation-section"
                            Visible="false">
                            <div class="validation-section-title warning">提醒 (請確認以下項目)</div>
                            <asp:BulletedList ID="blWarnings" runat="server"
                                CssClass="validation-list warning">
                            </asp:BulletedList>
                        </asp:Panel>
                    </div>
                    <div class="modalFooter">
                        <asp:Button ID="btnValidationBack" runat="server" Text="返回修改"
                            CssClass="btn btn-secondary" OnClick="btnValidationBack_Click" />
                        <asp:Button ID="btnValidationConfirm" runat="server" Text="確定仍要新增"
                            CssClass="btn btn-warning" OnClick="btnValidationConfirm_Click"
                            Visible="false" />
                    </div>
                </asp:Panel>
            </ContentTemplate>
            <Triggers>
                <asp:PostBackTrigger ControlID="btnUpload" />
            </Triggers>
        </asp:UpdatePanel>
    </form>
</body>

</html>
