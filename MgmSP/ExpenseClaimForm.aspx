<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ExpenseClaimForm.aspx.vb" Inherits="MgmSP.ExpenseClaimForm"
    MaintainScrollPositionOnPostback="true" %>
    <%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>

        <!DOCTYPE html>
        <html xmlns="http://www.w3.org/1999/xhtml">

        <head runat="server">
            <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
            <title>費用申請單</title>
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
                    
                    /* 強調色 */
                    --accent-primary: #3B4A6B;
                    --accent-hover: #4A5D82;
                    --accent-light: #7C8DB0;
                    
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

                /* Layout Grid - 保持原有配置 */
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
                select:focus,
                textarea:focus {
                    outline: none;
                    border-color: var(--accent-light);
                    box-shadow: 0 0 0 3px rgba(124, 141, 176, 0.12);
                }

                input[type="text"]::placeholder,
                textarea::placeholder {
                    color: var(--text-muted);
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
                    background: linear-gradient(135deg, var(--accent-hover) 0%, #5A6D92 100%);
                    box-shadow: 0 4px 12px rgba(59, 74, 107, 0.25);
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

                .btn-warning:hover {
                    box-shadow: 0 4px 12px rgba(201, 162, 39, 0.3);
                }

                .btn-info {
                    background: linear-gradient(135deg, var(--info) 0%, #6B8BA9 100%);
                    color: white;
                }

                .btn-info:hover {
                    box-shadow: 0 4px 12px rgba(91, 123, 154, 0.3);
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

                /* Tabs - 保持原有配置，更新視覺 */
                .tab-container {
                    display: flex;
                    border-bottom: 2px solid var(--accent-primary);
                    margin-top: 24px;
                }

                .tab-button {
                    padding: 10px 25px;
                    background-color: var(--bg-secondary);
                    border: 1px solid var(--border-color);
                    border-bottom: none;
                    cursor: pointer;
                    margin-right: 2px;
                    border-radius: 8px 8px 0 0;
                    font-weight: 500;
                    color: var(--text-secondary);
                    transition: all 0.2s ease;
                }

                .tab-button:hover {
                    background-color: var(--border-color);
                    color: var(--text-primary);
                }

                .tab-button.active {
                    background: linear-gradient(135deg, var(--accent-primary) 0%, var(--accent-hover) 100%);
                    color: white;
                    border-color: var(--accent-primary);
                }

                .tab-content {
                    display: none;
                    padding: 20px;
                    border: 1px solid var(--border-color);
                    border-top: none;
                    background-color: var(--bg-white);
                    min-height: 200px;
                    border-radius: 0 0 8px 8px;
                }

                .tab-content.active {
                    display: block;
                }

                /* GridView - 保持原有配置，更新視覺 */
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
                .gridview select {
                    width: 95%;
                    padding: 6px 8px;
                }

                /* Status Badges - 更新為優雅配色 */
                .badge {
                    padding: 5px 12px;
                    border-radius: 20px;
                    color: white;
                    font-size: 12px;
                    font-weight: 500;
                    letter-spacing: 0.03em;
                }

                .status-P {
                    background: linear-gradient(135deg, var(--text-secondary) 0%, #7A8A9B 100%);
                }

                /* Draft */
                .status-W {
                    background: linear-gradient(135deg, var(--warning) 0%, #D4AF37 100%);
                    color: white;
                }

                /* Pending */
                .status-A {
                    background: linear-gradient(135deg, var(--success) 0%, #7BA393 100%);
                }

                /* Approved */
                .status-R {
                    background: linear-gradient(135deg, var(--danger) 0%, #B86E68 100%);
                }

                /* Rejected */

                /* Modal - 保持原有配置，更新視覺 */
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

                /* ========================================
                   Header & Breadcrumb
                   ======================================== */
                .site-header {
                    background: linear-gradient(135deg, #1a1f2e 0%, #2a3142 100%);
                    padding: 1.25rem 2rem;
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
                }

                .site-logo span {
                    font-weight: 500;
                }

                .user-info {
                    color: var(--text-muted);
                    font-size: 0.875rem;
                }

                .user-info strong {
                    color: #E2E8F0;
                }

                .user-info a {
                    color: var(--text-muted);
                    text-decoration: none;
                    margin-left: 0.5rem;
                }

                .user-info a:hover {
                    color: #E2E8F0;
                }

                .breadcrumb {
                    padding: 1rem 2rem;
                    background: var(--bg-white);
                    border-bottom: 1px solid var(--border-color);
                    font-size: 0.875rem;
                    color: var(--text-secondary);
                }

                .breadcrumb a {
                    color: var(--accent-primary);
                    text-decoration: none;
                }

                .breadcrumb a:hover {
                    color: var(--accent-hover);
                }

                .breadcrumb span {
                    margin: 0 0.5rem;
                    color: var(--text-muted);
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

                /* 複製彈窗 */
                .copy-modal-overlay {
                    display: none;
                    position: fixed;
                    z-index: 1000;
                    left: 0;
                    top: 0;
                    width: 100%;
                    height: 100%;
                    background: rgba(26, 31, 46, 0.5);
                }

                .copy-modal-content {
                    background: var(--bg-white);
                    width: 420px;
                    margin: 12% auto;
                    padding: 24px;
                    border-radius: 12px;
                    box-shadow: var(--shadow-md);
                    border: 1px solid var(--border-color);
                }

                .copy-modal-title {
                    font-weight: 500;
                    margin-bottom: 12px;
                    color: var(--text-primary);
                }

                .copy-modal-question {
                    margin-bottom: 18px;
                    color: var(--text-secondary);
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

                /* ========================================
                   Site Header & Breadcrumb
                   ======================================== */
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
            </style>
            <script type="text/javascript">
                function switchTab(tabName) {
                    var hf = document.getElementById('<%= hfActiveTab.ClientID %>');
                    if (hf) hf.value = tabName;

                    // Remove active class
                    var tabBtns = document.getElementsByClassName('tab-button');
                    for (var i = 0; i < tabBtns.length; i++) {
                        tabBtns[i].className = tabBtns[i].className.replace(" active", "");
                    }

                    var tabContents = document.getElementsByClassName('tab-content');
                    for (var i = 0; i < tabContents.length; i++) {
                        tabContents[i].style.display = 'none';
                        tabContents[i].className = tabContents[i].className.replace(" active", "");
                    }

                    // Add active class
                    if (tabName === 'expense') {
                        document.getElementById('btnTabExpense').className += " active";
                        document.getElementById('divContentExpense').style.display = 'block';
                        document.getElementById('divContentExpense').className += " active";
                    } else if (tabName === 'mdr') {
                        document.getElementById('btnTabMDR').className += " active";
                        document.getElementById('divContentMDR').style.display = 'block';
                        document.getElementById('divContentMDR').className += " active";
                    }
                    return false;
                }

                function confirmDelete() {
                    return confirm('確定要刪除此筆費用申請單嗎？此操作無法復原。');
                }

                // 防止按鈕連續點擊
                var isSubmitting = false;

                function preventDoubleClick(btn, msg) {
                    // [Safety Check] 防止狀態卡死：如果 ASP.NET AJAX 沒在跑 PostBack，但 isSubmitting 還是 true，表示上次結束沒重置
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

                    // Client validation check
                    if (typeof (Page_ClientValidate) == 'function') {
                        if (Page_ClientValidate() == false) {
                            return false;
                        }
                    }

                    if (msg && !confirm(msg)) {
                        return false;
                    }

                    isSubmitting = true;
                    // 使用 setTimeout 確保按鈕在 PostBack 觸發後才禁用 (避免某些瀏覽器不送出 PostBack)
                    setTimeout(function () {
                        if (isSubmitting) { // 再次確認，防止同時被 reset
                            btn.disabled = true;
                            btn.value = '處理中...';
                        }
                    }, 50);
                    return true;
                }

                // 初始化設定 (確保 DOM 和 ScriptManager 載入後執行)
                function initPageLogic() {
                    if (typeof (Sys) !== 'undefined' && Sys.WebForms && Sys.WebForms.PageRequestManager) {
                        var prm = Sys.WebForms.PageRequestManager.getInstance();

                        // 移除舊 Handler (避免重複) - 雖然這個 API 沒有 removeAll，但我們只做一次

                        prm.add_endRequest(function () {
                            // 重置提交狀態
                            isSubmitting = false;

                            // 重新啟用按鈕 (確保即使 UpdatePanel 沒更新到按鈕，也能手動還原)
                            // 注意：如果 UpdatePanel 更新了按鈕，新按鈕預設就是啟用的，這裡主要處理沒更新的情況
                            var btns = document.querySelectorAll('.btn');
                            for (var i = 0; i < btns.length; i++) {
                                btns[i].disabled = false;
                                // 我們不還原 value，因為如果 PostBack 成功，UpdatePanel 應該已還原文字；
                                // 若失敗或部分更新，保留 '處理中' 也不一定是壞事，或者可以依需求還原。
                            }

                            // Keep tab state logic
                            var hf = document.getElementById('<%= hfActiveTab.ClientID %>');
                            if (hf && hf.value) {
                                switchTab(hf.value);
                            }
                        });

                        prm.add_pageLoaded(function () {
                            var hf = document.getElementById('<%= hfActiveTab.ClientID %>');
                            if (hf && hf.value) {
                                switchTab(hf.value);
                            }
                        });
                    }
                }

                // 註冊初始化 (支援 IE 和現代瀏覽器)
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
                    <asp:HyperLink ID="lnkExpenseModule" runat="server" NavigateUrl="~/ExpenseClaimList.aspx?smid=ec&smode=1">費用管理</asp:HyperLink>
                    <span class="separator">></span>
                    <span class="current">費用申請</span>
                </nav>

                <asp:ScriptManager ID="ScriptManager1" runat="server"></asp:ScriptManager>
                <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                    <ContentTemplate>
                        <asp:HiddenField ID="hfActiveTab" runat="server" Value="expense" />
                        <%-- [F] 用於儲存匯率日期，以便驗證時檢查 --%>
                            <asp:HiddenField ID="hfRateDate" runat="server" Value="" />
                            <div class="form-container">
                                <div
                                    style="display:flex; justify-content:space-between; align-items:center; margin-bottom:20px;">
                                    <h2 class="page-header-title">費用申請單 (Expense Claim)</h2>
                                    <div style="text-align:right;">
                                        <div style="margin-bottom:6px;">
                                            <asp:Button ID="btnCopyDocument" runat="server" Text="複製單據"
                                                OnClick="btnCopyDocument_Click" CssClass="btn btn-secondary" Visible="false"
                                                OnClientClick="return showCopyDialogForForm();" />
                                        </div>
                                        <asp:Label ID="lblDocNum" runat="server" Text="[New]" Font-Bold="True"
                                            Font-Size="18px" ForeColor="#3B4A6B"></asp:Label>
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
                                            <label class="form-label"><span class="required">*</span>供應商代碼:</label>
                                            <div class="form-control">
                                                <!-- Search Button Combo -->
                                                <div class="search-combo">
                                                    <asp:TextBox ID="txtCardCode" runat="server" placeholder="請點選搜尋"
                                                        ReadOnly="false">
                                                    </asp:TextBox>
                                                    <asp:Button ID="btnSearchCardCode" runat="server" Text="🔍"
                                                        CssClass="btn btn-secondary"
                                                        OnClick="btnSearchCardCode_Click" />
                                                </div>
                                                <div class="vendor-info">
                                                    <asp:Label ID="lblVendorInfo" runat="server">
                                                    </asp:Label>
                                                </div>
                                                <asp:Label ID="lblErrCardCode" runat="server" CssClass="error-text"
                                                    Visible="False"></asp:Label>
                                            </div>
                                        </div>
                                        <div class="form-group">
                                            <label class="form-label">供應商名稱:</label>
                                            <div class="form-control">
                                                <div class="search-combo">
                                                    <asp:TextBox ID="txtCardName" runat="server" placeholder="請點選搜尋"
                                                        ReadOnly="false">
                                                    </asp:TextBox>
                                                    <asp:Button ID="btnSearchCardName" runat="server" Text="🔍"
                                                        CssClass="btn btn-secondary"
                                                        OnClick="btnSearchCardName_Click" />
                                                </div>
                                                <asp:Label ID="lblErrCardName" runat="server" CssClass="error-text"
                                                    Visible="False"></asp:Label>
                                            </div>
                                        </div>
                                        <div class="form-group">
                                            <label class="form-label">供應商參考號:</label>
                                            <div class="form-control">
                                                <asp:TextBox ID="txtNumAtCard" runat="server"></asp:TextBox>
                                            </div>
                                        </div>
                                        <div class="form-group">
                                            <label class="form-label">文件幣別/匯率:</label>
                                            <div class="form-control">
                                                <asp:DropDownList ID="ddlDocCurrency" runat="server" AutoPostBack="True"
                                                    OnSelectedIndexChanged="ddlDocCurrency_SelectedIndexChanged"
                                                    Width="40%" style="margin-right:5px;"></asp:DropDownList>
                                                <asp:TextBox ID="txtDocRate" runat="server" Width="30%" Text="1.0"
                                                    AutoPostBack="true" OnTextChanged="txtDocRate_TextChanged">
                                                </asp:TextBox>
                                                <asp:Button ID="btnRefreshRate" runat="server" Text="↻"
                                                    CssClass="btn btn-secondary btn-icon" OnClick="btnRefreshRate_Click"
                                                    ToolTip="更新匯率" />
                                            </div>
                                        </div>
                                        <div class="form-group">
                                            <label class="form-label">收貨地址名稱:</label>
                                            <div class="form-control">
                                                <asp:DropDownList ID="ddlDeliveryAddr" runat="server"
                                                    AutoPostBack="true"
                                                    OnSelectedIndexChanged="ddlDeliveryAddr_SelectedIndexChanged">
                                                </asp:DropDownList>
                                            </div>
                                        </div>
                                        <div class="form-group">
                                            <label class="form-label">收貨地址:</label>
                                            <div class="form-control">
                                                <asp:TextBox ID="txtAddress" runat="server"></asp:TextBox>
                                            </div>
                                        </div>
                                        <div class="form-group">
                                            <label class="form-label">付款條件:</label>
                                            <div class="form-control">
                                                <asp:DropDownList ID="ddlGroupNum" runat="server" AutoPostBack="true"
                                                    OnSelectedIndexChanged="ddlGroupNum_SelectedIndexChanged"></asp:DropDownList>
                                            </div>
                                        </div>
                                        <div class="form-group">
                                            <label class="form-label">付款條件(列印):</label>
                                            <div class="form-control">
                                                <asp:TextBox ID="txtPymntGroup" runat="server" placeholder="列印用付款條件名稱">
                                                </asp:TextBox>
                                            </div>
                                        </div>
                                    </div>

                                    <!-- Right Column -->
                                    <div class="col-half">
                                        <div class="form-group">
                                            <label class="form-label">平台單號 (jID):</label>
                                            <div class="form-control">
                                                <asp:TextBox ID="txtJID" runat="server" ReadOnly="true"
                                                    CssClass="readonly-field" placeholder="系統自動產生"></asp:TextBox>
                                            </div>
                                        </div>
                                        <div class="form-group">
                                            <label class="form-label">AP單號 (B1):</label>
                                            <div class="form-control">
                                                <asp:TextBox ID="txtB1DocEntry" runat="server" ReadOnly="true"
                                                    CssClass="readonly-field" placeholder="SAP DocEntry"></asp:TextBox>
                                            </div>
                                        </div>
                                        <div class="form-group">
                                            <label class="form-label">簽核系統 PID:</label>
                                            <div class="form-control">
                                                <asp:TextBox ID="txtUPID" runat="server"></asp:TextBox>
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
                                            <label class="form-label"><span class="required">*</span>過帳日期:</label>
                                            <div class="form-control">
                                                <asp:TextBox ID="txtDocDate" runat="server" TextMode="Date"
                                                    AutoPostBack="true" OnTextChanged="btnRefreshRate_Click">
                                                </asp:TextBox>
                                                <asp:Label ID="lblErrDocDate" runat="server" CssClass="error-text"
                                                    Visible="False"></asp:Label>
                                            </div>
                                        </div>
                                        <div class="form-group">
                                            <label class="form-label"><span class="required">*</span>到期日:</label>
                                            <div class="form-control">
                                                <asp:TextBox ID="txtDocDueDate" runat="server" TextMode="Date">
                                                </asp:TextBox>
                                                <asp:Label ID="lblErrDocDueDate" runat="server" CssClass="error-text"
                                                    Visible="False"></asp:Label>
                                            </div>
                                        </div>
                                        <div class="form-group">
                                            <label class="form-label"><span class="required">*</span>文件日期:</label>
                                            <div class="form-control">
                                                <asp:TextBox ID="txtTaxDate" runat="server" TextMode="Date">
                                                </asp:TextBox>
                                                <asp:Label ID="lblErrTaxDate" runat="server" CssClass="error-text"
                                                    Visible="False"></asp:Label>
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
                                            <label class="form-label">核准人:</label>
                                            <div class="form-control">
                                                <asp:TextBox ID="txtApprovedBy" runat="server" ReadOnly="true"
                                                    CssClass="readonly-field"></asp:TextBox>
                                            </div>
                                        </div>
                                    </div>
                                </div>

                                <!-- Tabs -->
                                <div class="tab-container">
                                    <button type="button" class="tab-button active" id="btnTabExpense" runat="server"
                                        onclick="switchTab('expense'); return false;">費用申請明細</button>
                                    <button type="button" class="tab-button" id="btnTabMDR" runat="server"
                                        onclick="switchTab('mdr'); return false;">憑證明細</button>
                                </div>

                                <!-- Tab 1: Expense Lines -->
                                <div id="divContentExpense" class="tab-content active" runat="server"
                                    ClientIDMode="Static">
                                    <div style="margin-bottom: 10px; display:flex; justify-content:space-between;">
                                        <div>
                                            <asp:Button ID="btnAddLine" runat="server" Text="+ 新增明細"
                                                OnClick="btnAddLine_Click" CssClass="btn btn-primary" />
                                            <asp:Button ID="btnDeleteLine" runat="server" Text="🗑 刪除選取"
                                                OnClick="btnDeleteLine_Click" CssClass="btn btn-danger"
                                                OnClientClick="return confirm('確定刪除選中的明細行？');" />
                                            <asp:Button ID="btnGenerateMDR" runat="server" Text="📋 產生憑證明細"
                                                OnClick="btnGenerateMDR_Click" CssClass="btn btn-warning"
                                                ToolTip="依據費用明細自動產生憑證明細" />
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
                                                        <%-- [C] 使用 Handler 安全下載附件，避免路徑洩漏 --%>
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
                                        <asp:GridView ID="gvExpenseDetail" runat="server" AutoGenerateColumns="False"
                                            CssClass="gridview" OnRowDataBound="gvExpenseDetail_RowDataBound"
                                            OnRowCommand="gvExpenseDetail_RowCommand">
                                            <Columns>
                                                <asp:TemplateField HeaderText="選">
                                                    <HeaderTemplate>
                                                        <input type="checkbox"
                                                            onclick="toggleGridCheckboxes(this, '<%= gvExpenseDetail.ClientID %>', 'chkSelect')" />
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

                                                <%-- 幣別與匯率欄位已隱藏，寫入時使用單頭的幣別與匯率 --%>

                                                    <asp:TemplateField HeaderText="費用類別">
                                                        <ItemTemplate>
                                                            <asp:DropDownList ID="ddlExpCategory" runat="server"
                                                                Width="150px" AutoPostBack="true"
                                                                OnSelectedIndexChanged="ddlExpCategory_SelectedIndexChanged">
                                                            </asp:DropDownList>
                                                        </ItemTemplate>
                                                        <ItemStyle Width="160px" />
                                                    </asp:TemplateField>

                                                    <asp:TemplateField HeaderText="說明">
                                                        <ItemTemplate>
                                                            <asp:TextBox ID="txtDescription" runat="server"
                                                                Width="200px" AutoPostBack="true"
                                                                OnTextChanged="txtDescription_TextChanged">
                                                            </asp:TextBox>
                                                        </ItemTemplate>
                                                        <ItemStyle Width="210px" />
                                                    </asp:TemplateField>

                                                    <asp:TemplateField HeaderText="會計科目">
                                                        <ItemTemplate>
                                                            <div style="display:flex; align-items:center; gap:4px;">
                                                                <asp:TextBox ID="txtAcctCode" runat="server" Width="80px"
                                                                    ReadOnly="true" CssClass="readonly-field"></asp:TextBox>
                                                                <asp:Button ID="btnSearchAcct" runat="server" Text="🔍"
                                                                    CssClass="btn btn-secondary btn-icon"
                                                                    CommandName="SearchAcct"
                                                                    CommandArgument='<%# Container.DataItemIndex %>' />
                                                            </div>
                                                        </ItemTemplate>
                                                        <ItemStyle Width="90px" />
                                                    </asp:TemplateField>

                                                    <asp:TemplateField HeaderText="未稅金額">
                                                        <ItemTemplate>
                                                            <asp:TextBox ID="txtLineTotal" runat="server" Width="90px"
                                                                style="text-align:right;" AutoPostBack="true"
                                                                OnTextChanged="CalculateLineTotal"></asp:TextBox>
                                                        </ItemTemplate>
                                                        <ItemStyle Width="100px" />
                                                    </asp:TemplateField>

                                                    <asp:TemplateField HeaderText="稅別">
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
                                                                Width="70px" style="text-align:right;"
                                                                AutoPostBack="true" OnTextChanged="CalculateVatSum">
                                                            </asp:TextBox>
                                                        </ItemTemplate>
                                                        <ItemStyle Width="80px" />
                                                    </asp:TemplateField>

                                                    <asp:TemplateField HeaderText="含稅金額">
                                                        <ItemTemplate>
                                                            <asp:TextBox ID="txtPriceAfterVat" runat="server"
                                                                Width="90px" style="text-align:right;"
                                                                AutoPostBack="true"
                                                                OnTextChanged="CalculatePriceAfterVat"></asp:TextBox>
                                                        </ItemTemplate>
                                                        <ItemStyle Width="100px" />
                                                    </asp:TemplateField>

                                                    <asp:TemplateField HeaderText="產品">
                                                        <ItemTemplate>
                                                            <asp:DropDownList ID="ddlCostingCode" runat="server"
                                                                Width="100px"></asp:DropDownList>
                                                        </ItemTemplate>
                                                        <ItemStyle Width="110px" />
                                                    </asp:TemplateField>

                                                    <asp:TemplateField HeaderText="部門">
                                                        <ItemTemplate>
                                                            <asp:DropDownList ID="ddlCostingCode2" runat="server"
                                                                Width="100px" AutoPostBack="true"
                                                                OnSelectedIndexChanged="ddlCostingCode2_SelectedIndexChanged"></asp:DropDownList>
                                                        </ItemTemplate>
                                                        <ItemStyle Width="110px" />
                                                    </asp:TemplateField>
                                            </Columns>
                                            <EmptyDataTemplate>
                                                <div class="empty-data-hint">請新增費用明細</div>
                                            </EmptyDataTemplate>
                                        </asp:GridView>
                                    </div>
                                </div>

                                <!-- Tab 2: MDR Invoice Details -->
                                <div id="divContentMDR" class="tab-content" runat="server" ClientIDMode="Static">
                                    <div style="margin-bottom: 10px;">
                                        <asp:Button ID="btnAddMDRRow" runat="server" Text="+ 新增憑證"
                                            OnClick="btnAddMDRRow_Click" CssClass="btn btn-primary" />
                                        <asp:Button ID="btnDeleteMDRRow" runat="server" Text="🗑 刪除選取"
                                            OnClick="btnDeleteMDRRow_Click" CssClass="btn btn-danger" />
                                    </div>

                                    <div style="overflow-x:auto;">
                                        <asp:GridView ID="gvMDRDetail" runat="server" AutoGenerateColumns="False"
                                            CssClass="gridview" OnRowDataBound="gvMDRDetail_RowDataBound">
                                            <Columns>
                                                <asp:TemplateField HeaderText="選">
                                                    <HeaderTemplate>
                                                        <input type="checkbox"
                                                            onclick="toggleGridCheckboxes(this, '<%= gvMDRDetail.ClientID %>', 'chkSelectMDR')" />
                                                    </HeaderTemplate>
                                                    <ItemTemplate>
                                                        <asp:CheckBox ID="chkSelectMDR" runat="server" />
                                                    </ItemTemplate>
                                                    <ItemStyle Width="30px" HorizontalAlign="Center" />
                                                </asp:TemplateField>

                                                <%-- 供應商代碼欄位已移除，寫入時使用單頭供應商 --%>

                                                    <asp:TemplateField HeaderText="統一編號">
                                                        <ItemTemplate>
                                                            <asp:TextBox ID="txtSTCEG" runat="server"
                                                                Text='<%# Bind("U_STCEG") %>' Width="90px">
                                                            </asp:TextBox>
                                                        </ItemTemplate>
                                                        <ItemStyle Width="100px" />
                                                    </asp:TemplateField>

                                                    <asp:TemplateField HeaderText="憑證號碼">
                                                        <ItemTemplate>
                                                            <asp:TextBox ID="txtXBLNR" runat="server"
                                                                Text='<%# Bind("U_XBLNR") %>' Width="110px"
                                                                AutoPostBack="true"
                                                                OnTextChanged="txtXBLNR_TextChanged">
                                                            </asp:TextBox>
                                                        </ItemTemplate>
                                                        <ItemStyle Width="120px" />
                                                    </asp:TemplateField>

                                                    <asp:TemplateField HeaderText="憑證類型">
                                                        <ItemTemplate>
                                                            <asp:DropDownList ID="ddlZFORM_CODE" runat="server"
                                                                SelectedValue='<%# Bind("U_ZFORM_CODE") %>'
                                                                Width="200px">
                                                                <asp:ListItem Value="21" Text="21-三聯手開發票">
                                                                </asp:ListItem>
                                                                <asp:ListItem Value="22" Text="22-高鐵/二聯收銀機（長條）">
                                                                </asp:ListItem>
                                                                <asp:ListItem Value="25" Text="25-電子發票/公營事業/三聯收銀機">
                                                                </asp:ListItem>
                                                                <asp:ListItem Value="28" Text="28-海關代徵營業稅">
                                                                </asp:ListItem>
                                                                <asp:ListItem Value="99" Text="99-其他"></asp:ListItem>
                                                            </asp:DropDownList>

                                                        </ItemTemplate>
                                                        <ItemStyle Width="160px" />
                                                    </asp:TemplateField>

                                                    <asp:TemplateField HeaderText="憑證日期">
                                                        <ItemTemplate>
                                                            <asp:TextBox ID="txtBLDAT" runat="server"
                                                                Text='<%# Bind("U_BLDAT", "{0:yyyy-MM-dd}") %>'
                                                                TextMode="Date" Width="120px"></asp:TextBox>
                                                        </ItemTemplate>
                                                        <ItemStyle Width="130px" />
                                                    </asp:TemplateField>

                                                    <asp:TemplateField HeaderText="營業稅日期">
                                                        <ItemTemplate>
                                                            <asp:TextBox ID="txtVATDATE" runat="server"
                                                                Text='<%# Bind("U_VATDATE", "{0:yyyy-MM-dd}") %>'
                                                                TextMode="Date" Width="120px"></asp:TextBox>
                                                        </ItemTemplate>
                                                        <ItemStyle Width="130px" />
                                                    </asp:TemplateField>

                                                    <asp:TemplateField HeaderText="未稅金額">
                                                        <ItemTemplate>
                                                            <asp:TextBox ID="txtHWBAS" runat="server"
                                                                Text='<%# Bind("U_HWBAS", "{0:N2}") %>' Width="90px"
                                                                style="text-align:right;" AutoPostBack="true"
                                                                OnTextChanged="CalculateMDRTotal"></asp:TextBox>
                                                        </ItemTemplate>
                                                        <ItemStyle Width="100px" />
                                                    </asp:TemplateField>

                                                    <asp:TemplateField HeaderText="稅別">
                                                        <ItemTemplate>
                                                            <asp:DropDownList ID="ddlTAX_TYPE" runat="server"
                                                                SelectedValue='<%# Bind("U_TAX_TYPE") %>'
                                                                AutoPostBack="true"
                                                                OnSelectedIndexChanged="CalculateMDRTotal" Width="80px">
                                                                <asp:ListItem Value="1" Text="1-應稅"></asp:ListItem>
                                                                <asp:ListItem Value="2" Text="2-零稅"></asp:ListItem>
                                                                <asp:ListItem Value="3" Text="3-免稅"></asp:ListItem>
                                                            </asp:DropDownList>
                                                        </ItemTemplate>
                                                        <ItemStyle Width="90px" />
                                                    </asp:TemplateField>

                                                    <asp:TemplateField HeaderText="稅額">
                                                        <ItemTemplate>
                                                            <asp:TextBox ID="txtHWSTE" runat="server"
                                                                Text='<%# Bind("U_HWSTE", "{0:N2}") %>' Width="80px"
                                                                style="text-align:right;" AutoPostBack="true"
                                                                OnTextChanged="CalculateMDRTaxManual"></asp:TextBox>
                                                        </ItemTemplate>
                                                        <ItemStyle Width="90px" />
                                                    </asp:TemplateField>
                                            </Columns>
                                            <EmptyDataTemplate>
                                                <div class="empty-data-hint">
                                                    請新增憑證明細，或點擊「產生憑證明細」按鈕自動產生</div>
                                            </EmptyDataTemplate>
                                        </asp:GridView>
                                    </div>
                                </div>

                                <!-- Footer Section -->
                                <div class="footer-section">
                                    <div class="row">
                                        <div class="col-half">
                                            <div class="form-group">
                                                <label class="form-label">採購人員:</label>
                                                <div class="form-control">
                                                    <asp:DropDownList ID="ddlPurchaser" runat="server">
                                                    </asp:DropDownList>
                                                </div>
                                            </div>
                                            <div class="form-group">
                                                <label class="form-label">所有人:</label>
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
                                        <asp:Button ID="btnUpdateComment" runat="server" Text="發送意見 (Comment)"
                                            OnClick="btnUpdateComment_Click" CssClass="btn btn-warning" />
                                        <asp:Button ID="btnReject" runat="server" Text="退回 (Reject)"
                                            OnClick="btnReject_Click" CssClass="btn btn-danger"
                                            OnClientClick="return confirm('確定要退回此單據嗎？');" />
                                    </div>
                                </asp:Panel>

                                <!-- Buttons -->
                                <div style="text-align:center; margin-top:30px;">
                                    <!-- 新增/編輯模式按鈕 -->
                                    <asp:Button ID="btnSave" runat="server" Text="暫存 (Draft)" OnClick="btnSave_Click"
                                        CssClass="btn btn-primary" Visible="false" />
                                    <asp:Button ID="btnSubmit" runat="server" Text="儲存並送審 (Save & Submit)"
                                        OnClick="btnSubmit_Click" CssClass="btn btn-success"
                                        OnClientClick="return preventDoubleClick(this, '確定要送出審核嗎？');" />
                                    <asp:Button ID="btnDelete" runat="server" Text="刪除 (Delete)"
                                        OnClick="btnDelete_Click" CssClass="btn btn-danger"
                                        OnClientClick="return preventDoubleClick(this, null) && confirmDelete();" />
                                    <asp:Button ID="btnCancel" runat="server" Text="取消 (Cancel)"
                                        OnClick="btnCancel_Click" CssClass="btn btn-secondary" />

                                    <!-- 檢視模式按鈕 -->
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
                                                placeholder="輸入關鍵字...">
                                            </asp:TextBox>
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

                            <!-- Account Search Modal -->
                            <asp:Button ID="btnAcctDummy" runat="server" style="display:none" />
                            <ajaxToolkit:ModalPopupExtender ID="mpeAcct" runat="server"
                                TargetControlID="btnAcctDummy"
                                PopupControlID="pnlAcctSearch" BackgroundCssClass="modalBackground"
                                CancelControlID="btnCloseAcct" />
                            <asp:Panel ID="pnlAcctSearch" runat="server" CssClass="modalPopup" style="display:none;">
                                <div class="modalHeader">
                                    <span>會計科目搜尋</span>
                                    <asp:LinkButton ID="btnCloseAcct" runat="server" ForeColor="White"
                                        Font-Bold="true" style="text-decoration:none;">✕</asp:LinkButton>
                                </div>
                                <div class="modalBody">
                                    <div style="margin-bottom:10px;">
                                        <div style="display:flex; align-items:center;">
                                            <asp:TextBox ID="txtAcctSearchKeyword" runat="server"
                                                placeholder="輸入會計科目代碼或名稱...">
                                            </asp:TextBox>
                                            <asp:Button ID="btnDoSearchAcct" runat="server" Text="搜尋"
                                                OnClick="btnDoSearchAcct_Click" CssClass="btn btn-primary"
                                                style="margin-left:5px;" />
                                            <asp:HiddenField ID="hfAcctSearchRowIndex" runat="server" />
                                        </div>
                                        <div style="margin-top:8px;">
                                                <asp:RadioButtonList ID="rblAcctSearchMode" runat="server"
                                                    RepeatDirection="Horizontal">
                                                    <asp:ListItem Value="Fuzzy">模糊搜尋</asp:ListItem>
                                                    <asp:ListItem Value="Exact" Selected="True">開頭比對</asp:ListItem>
                                                </asp:RadioButtonList>
                                        </div>
                                    </div>
                                    <asp:GridView ID="gvAcctSearch" runat="server" AutoGenerateColumns="False"
                                        Width="100%" CssClass="gridview" OnRowCommand="gvAcctSearch_RowCommand"
                                        AllowPaging="True" PageSize="10"
                                        AllowSorting="True"
                                        OnSorting="gvAcctSearch_Sorting"
                                        OnPageIndexChanging="gvAcctSearch_PageIndexChanging">
                                        <Columns>
                                            <asp:TemplateField HeaderText="動作">
                                                <ItemTemplate>
                                                    <asp:LinkButton ID="lbtnSelectAcct" runat="server"
                                                        CommandName="SelectAcct"
                                                        CommandArgument='<%# Eval("AcctCode") + "|" + Eval("AcctName") %>'
                                                        CssClass="btn btn-success btn-icon">選取</asp:LinkButton>
                                                </ItemTemplate>
                                                <ItemStyle HorizontalAlign="Center" Width="70px" />
                                            </asp:TemplateField>
                                            <asp:BoundField DataField="AcctCode" HeaderText="代碼" SortExpression="AcctCode" />
                                            <asp:BoundField DataField="AcctName" HeaderText="名稱" SortExpression="AcctName" />
                                        </Columns>
                                        <PagerStyle HorizontalAlign="Center" CssClass="gridview" />
                                    </asp:GridView>
                                </div>
                            </asp:Panel>

                            <asp:HiddenField ID="hfAcctPendingRowIndex" runat="server" />
                            <asp:Button ID="btnAcctRowLeave" runat="server" style="display:none"
                                OnClick="btnAcctRowLeave_Click" />

                            <script type="text/javascript">
                                (function () {
                                    var acctLeavePosting = false;

                                    function isChildOf(parent, node) {
                                        while (node) {
                                            if (node === parent) {
                                                return true;
                                            }
                                            node = node.parentNode;
                                        }
                                        return false;
                                    }

                                    function wireAcctRowLeave() {
                                        var grid = document.getElementById('<%= gvExpenseDetail.ClientID %>');
                                        if (!grid) {
                                            return;
                                        }
                                        var rows = grid.getElementsByTagName('tr');
                                        for (var i = 0; i < rows.length; i++) {
                                            var row = rows[i];
                                            if (!row.getAttribute('data-rowindex')) {
                                                continue;
                                            }
                                            row.addEventListener('focusout', function (e) {
                                                if (acctLeavePosting) {
                                                    return;
                                                }
                                                var pending = document.getElementById('<%= hfAcctPendingRowIndex.ClientID %>');
                                                if (!pending || pending.value === '') {
                                                    return;
                                                }
                                                var rowIndex = this.getAttribute('data-rowindex');
                                                if (rowIndex !== pending.value) {
                                                    return;
                                                }
                                                var related = e.relatedTarget;
                                                if (related && isChildOf(this, related)) {
                                                    return;
                                                }
                                                acctLeavePosting = true;
                                                __doPostBack('<%= btnAcctRowLeave.UniqueID %>', '');
                                            }, true);
                                        }
                                    }

                                    if (window.Sys && Sys.Application) {
                                        Sys.Application.add_load(wireAcctRowLeave);
                                    } else {
                                        window.addEventListener('load', wireAcctRowLeave);
                                    }
                                })();
                            </script>

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
                                    <!-- 錯誤區塊 -->
                                    <asp:Panel ID="pnlErrors" runat="server" CssClass="validation-section"
                                        Visible="false">
                                        <div class="validation-section-title error">錯誤 (必須修正才能儲存)</div>
                                        <asp:BulletedList ID="blErrors" runat="server" CssClass="validation-list error">
                                        </asp:BulletedList>
                                    </asp:Panel>
                                    <!-- 警告區塊 -->
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

                            <!-- 費用部門選擇彈窗 -->
                            <asp:Button ID="btnExpDeptDummy" runat="server" Style="display:none" />
                            <ajaxToolkit:ModalPopupExtender ID="mpeExpDept" runat="server"
                                BehaviorID="mpeExpDeptBehavior" TargetControlID="btnExpDeptDummy"
                                PopupControlID="pnlExpDept" BackgroundCssClass="modalBackground"
                                DropShadow="false" />
                            <asp:Panel ID="pnlExpDept" runat="server" CssClass="modalPopup" Style="display:none; width:400px;">
                                <div class="modalHeader" style="background: linear-gradient(135deg, #5B7B9A 0%, #6B8BA9 100%);">
                                    <span>選擇費用部門</span>
                                </div>
                                <div class="modalBody">
                                    <p style="margin-bottom:15px; color: var(--text-secondary);">您尚未設定費用部門，請選擇您所屬的費用部門：</p>
                                    <div class="form-group">
                                        <label class="form-label" style="width:100px;">費用部門:</label>
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

                            <asp:HiddenField ID="hfCopyAttachment" runat="server" Value="0" />
                            <asp:HiddenField ID="hfCopyMDR" runat="server" Value="0" />

                            <div id="divCopyModalForm" class="copy-modal-overlay">
                                <div class="copy-modal-content">
                                    <div class="copy-modal-title">複製選項</div>
                                    <div id="divCopyQuestionForm" class="copy-modal-question">是否複製附件？</div>
                                    <div style="text-align:right;">
                                        <button type="button" class="btn btn-success" onclick="copyFormDialogAnswer('yes');">是</button>
                                        <button type="button" class="btn btn-secondary" onclick="copyFormDialogAnswer('no');">否</button>
                                        <button type="button" class="btn btn-secondary" onclick="copyFormDialogAnswer('cancel');">取消</button>
                                    </div>
                                </div>
                            </div>

                            <script type="text/javascript">
                                var copyFormStage = '';
                                function showCopyDialogForForm() {
                                    copyFormStage = 'attach';
                                    var question = document.getElementById('divCopyQuestionForm');
                                    if (question) {
                                        question.innerText = '是否複製附件？';
                                    }
                                    var modal = document.getElementById('divCopyModalForm');
                                    if (modal) {
                                        modal.style.display = 'block';
                                    }
                                    return false;
                                }

                                function copyFormDialogAnswer(answer) {
                                    if (answer === 'cancel') {
                                        hideCopyDialogForForm();
                                        return;
                                    }
                                    if (copyFormStage === 'attach') {
                                        var hfAttach = document.getElementById('<%= hfCopyAttachment.ClientID %>');
                                        if (hfAttach) {
                                            hfAttach.value = (answer === 'yes') ? '1' : '0';
                                        }
                                        copyFormStage = 'mdr';
                                        var question = document.getElementById('divCopyQuestionForm');
                                        if (question) {
                                            question.innerText = '是否複製憑證明細？';
                                        }
                                        return;
                                    }
                                    if (copyFormStage === 'mdr') {
                                        var hfMdr = document.getElementById('<%= hfCopyMDR.ClientID %>');
                                        if (hfMdr) {
                                            hfMdr.value = (answer === 'yes') ? '1' : '0';
                                        }
                                        hideCopyDialogForForm();
                                        __doPostBack('<%= btnCopyDocument.UniqueID %>', '');
                                    }
                                }

                                function hideCopyDialogForForm() {
                                    var modal = document.getElementById('divCopyModalForm');
                                    if (modal) {
                                        modal.style.display = 'none';
                                    }
                                }
                            </script>
                    </ContentTemplate>
                    <Triggers>
                        <asp:PostBackTrigger ControlID="btnUpload" />
                    </Triggers>
                </asp:UpdatePanel>
            </form>
        </body>

        </html>

