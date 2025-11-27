<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ExpenseClaimForm.aspx.vb"
    Inherits="MgmSP.ExpenseClaimForm" MaintainScrollPositionOnPostback="true" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>費用申請單</title>
    <style type="text/css">
        body { font-family: "Microsoft JhengHei", Arial, sans-serif; font-size: 14px; background-color: #f5f5f5; }
        .form-container { max-width: 1400px; margin: 20px auto; padding: 20px; background-color: white; box-shadow: 0 0 10px rgba(0,0,0,0.1); border-radius: 5px; }
        
        .section-header { 
            background-color: #4CAF50; color: white; padding: 10px 15px;
            margin: 20px 0 15px 0; font-weight: bold; border-radius: 4px; font-size: 16px;
        }
        
        /* Layout Grid */
        .row { display: flex; flex-wrap: wrap; margin-right: -15px; margin-left: -15px; }
        .col-half { flex: 0 0 50%; max-width: 50%; padding-right: 15px; padding-left: 15px; box-sizing: border-box; }
        
        .form-group { margin-bottom: 12px; display: flex; align-items: center; }
        .form-label { width: 140px; font-weight: bold; padding-right: 10px; text-align: right; color: #333; }
        .form-control { flex: 1; display: flex; align-items: center; }
        
        input[type="text"], input[type="date"], select, textarea {
            padding: 6px 10px; border: 1px solid #ccc; border-radius: 4px;
            font-family: inherit; font-size: 14px; width: 100%; box-sizing: border-box;
        }
        textarea { resize: vertical; }
        
        .readonly-field { background-color: #e9ecef; cursor: not-allowed; }
        
        .btn { padding: 8px 20px; border-radius: 4px; border: none; cursor: pointer; font-size: 14px; font-weight: bold; margin-right: 5px; }
        .btn-primary { background-color: #007bff; color: white; }
        .btn-success { background-color: #28a745; color: white; }
        .btn-danger { background-color: #dc3545; color: white; }
        .btn-secondary { background-color: #6c757d; color: white; }
        .btn-warning { background-color: #ffc107; color: black; }
        .btn:hover { opacity: 0.9; }
        
        .btn-icon { padding: 4px 10px; font-size: 12px; margin-left: 5px; }

        .required { color: red; margin-right: 3px; }
        .error-text { color: #dc3545; font-size: 12px; margin-left: 5px; display: block; }
        
        /* Tabs */
        .tab-container { display: flex; border-bottom: 2px solid #4CAF50; margin-top: 20px; }
        .tab-button { 
            padding: 10px 25px; background-color: #f8f9fa; border: 1px solid #dee2e6; border-bottom: none;
            cursor: pointer; margin-right: 2px; border-radius: 4px 4px 0 0; font-weight: bold; color: #495057;
        }
        .tab-button.active { background-color: #4CAF50; color: white; border-color: #4CAF50; }
        .tab-content { display: none; padding: 20px; border: 1px solid #dee2e6; border-top: none; background-color: white; min-height: 200px; }
        .tab-content.active { display: block; }
        
        /* GridView */
        .gridview { border-collapse: collapse; width: 100%; margin-top: 10px; font-size: 13px; }
        .gridview th { background-color: #f1f1f1; color: #333; padding: 10px; border: 1px solid #ddd; text-align: center; white-space: nowrap; }
        .gridview td { padding: 5px; border: 1px solid #ddd; vertical-align: middle; }
        .gridview input[type="text"], .gridview select { width: 95%; padding: 4px; }
        
        /* Status Badges */
        .badge { padding: 5px 10px; border-radius: 10px; color: white; font-size: 12px; font-weight: bold; }
        .status-P { background-color: #6c757d; } /* Draft */
        .status-W { background-color: #ffc107; color: black; } /* Pending */
        .status-A { background-color: #28a745; } /* Approved */
        .status-R { background-color: #dc3545; } /* Rejected */

        /* Modal */
        .modalBackground { background-color: rgba(0,0,0,0.5); }
        .modalPopup { background-color: white; border-radius: 5px; padding: 0; width: 700px; box-shadow: 0 5px 15px rgba(0,0,0,0.3); }
        .modalHeader { background-color: #4CAF50; color: white; padding: 10px 15px; border-radius: 5px 5px 0 0; font-weight: bold; display: flex; justify-content: space-between; align-items: center; }
        .modalBody { padding: 15px; max-height: 500px; overflow-y: auto; }
        .modalFooter { padding: 10px 15px; border-top: 1px solid #eee; text-align: right; }
    </style>
    <script type="text/javascript">
        function switchTab(tabName) {
            var hf = document.getElementById('<%= hfActiveTab.ClientID %>');
            if (hf) hf.value = tabName;

            // Remove active class
            document.querySelectorAll('.tab-button').forEach(btn => btn.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(div => div.classList.remove('active'));

            // Add active class
            if (tabName === 'expense') {
                document.getElementById('btnTabExpense').classList.add('active');
                document.getElementById('divContentExpense').classList.add('active');
            } else if (tabName === 'mdr') {
                document.getElementById('btnTabMDR').classList.add('active');
                document.getElementById('divContentMDR').classList.add('active');
            }
            return false;
        }

        // Keep tab state after postback
        Sys.WebForms.PageRequestManager.getInstance().add_endRequest(function () {
            var hf = document.getElementById('<%= hfActiveTab.ClientID %>');
            if (hf && hf.value) {
                switchTab(hf.value);
            }
        });

        function confirmDelete() {
            return confirm('確定要刪除此筆費用申請單嗎？此操作無法復原。');
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server"></asp:ScriptManager>
        <asp:HiddenField ID="hfActiveTab" runat="server" Value="expense" />

        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
            <ContentTemplate>
                <div class="form-container">
                    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:20px;">
                        <h2 style="margin:0; color:#4CAF50;">費用申請單 (Expense Claim)</h2>
                        <div style="text-align:right;">
                            <asp:Label ID="lblDocNum" runat="server" Text="[New]" Font-Bold="True" Font-Size="18px" ForeColor="#007bff"></asp:Label>
                            <br />
                            <asp:Label ID="lblDocStatus" runat="server" CssClass="badge status-P" Text="草稿"></asp:Label>
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
                                    <div style="display:flex; width:100%;">
                                        <asp:TextBox ID="txtCardCode" runat="server" placeholder="請點選搜尋" ReadOnly="false" style="border-top-right-radius:0; border-bottom-right-radius:0;"></asp:TextBox>
                                        <asp:Button ID="btnSearchCardCode" runat="server" Text="🔍" CssClass="btn btn-secondary" style="border-top-left-radius:0; border-bottom-left-radius:0; margin:0;" OnClick="btnSearchCardCode_Click" />
                                    </div>
                                    <div style="margin-top:5px; font-size:12px;">
                                        <asp:Label ID="lblVendorInfo" runat="server" ForeColor="Blue"></asp:Label>
                                    </div>
                                    <asp:Label ID="lblErrCardCode" runat="server" CssClass="error-text" Visible="False"></asp:Label>
                                </div>
                            </div>
                            <div class="form-group">
                                <label class="form-label">供應商名稱:</label>
                                <div class="form-control">
                                    <div style="display:flex; width:100%;">
                                        <asp:TextBox ID="txtCardName" runat="server" placeholder="請點選搜尋" ReadOnly="false" style="border-top-right-radius:0; border-bottom-right-radius:0;"></asp:TextBox>
                                        <asp:Button ID="btnSearchCardName" runat="server" Text="🔍" CssClass="btn btn-secondary" style="border-top-left-radius:0; border-bottom-left-radius:0; margin:0;" OnClick="btnSearchCardName_Click" />
                                    </div>
                                    <asp:Label ID="lblErrCardName" runat="server" CssClass="error-text" Visible="False"></asp:Label>
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
                                    <asp:DropDownList ID="ddlDocCurrency" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlDocCurrency_SelectedIndexChanged" Width="40%" style="margin-right:5px;"></asp:DropDownList>
                                    <asp:TextBox ID="txtDocRate" runat="server" Width="30%" Text="1.0"></asp:TextBox>
                                    <asp:Button ID="btnRefreshRate" runat="server" Text="↻" CssClass="btn btn-secondary btn-icon" OnClick="btnRefreshRate_Click" ToolTip="更新匯率" />
                                </div>
                            </div>
                            <div class="form-group">
                                <label class="form-label">收貨地址名稱:</label>
                                <div class="form-control">
                                    <asp:DropDownList ID="ddlDeliveryAddr" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlDeliveryAddr_SelectedIndexChanged"></asp:DropDownList>
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
                                    <asp:DropDownList ID="ddlGroupNum" runat="server"></asp:DropDownList>
                                </div>
                            </div>
                        </div>

                        <!-- Right Column -->
                        <div class="col-half">
                            <div class="form-group">
                                <label class="form-label">平台單號 (jID):</label>
                                <div class="form-control">
                                    <asp:TextBox ID="txtJID" runat="server" ReadOnly="true" CssClass="readonly-field" placeholder="系統自動產生"></asp:TextBox>
                                </div>
                            </div>
                            <div class="form-group">
                                <label class="form-label">AP單號 (B1):</label>
                                <div class="form-control">
                                    <asp:TextBox ID="txtB1DocEntry" runat="server" ReadOnly="true" CssClass="readonly-field" placeholder="SAP DocEntry"></asp:TextBox>
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
                                    <asp:TextBox ID="txtStatusDisplay" runat="server" ReadOnly="true" CssClass="readonly-field"></asp:TextBox>
                                </div>
                            </div>
                            <div class="form-group">
                                <label class="form-label">過帳日期 (Tax):</label>
                                <div class="form-control">
                                    <asp:TextBox ID="txtTaxDate" runat="server" TextMode="Date"></asp:TextBox>
                                </div>
                            </div>
                            <div class="form-group">
                                <label class="form-label"><span class="required">*</span>到期日:</label>
                                <div class="form-control">
                                    <asp:TextBox ID="txtDocDueDate" runat="server" TextMode="Date"></asp:TextBox>
                                    <asp:Label ID="lblErrDocDueDate" runat="server" CssClass="error-text" Visible="False"></asp:Label>
                                </div>
                            </div>
                            <div class="form-group">
                                <label class="form-label"><span class="required">*</span>文件日期:</label>
                                <div class="form-control">
                                    <asp:TextBox ID="txtDocDate" runat="server" TextMode="Date"></asp:TextBox>
                                    <asp:Label ID="lblErrDocDate" runat="server" CssClass="error-text" Visible="False"></asp:Label>
                                </div>
                            </div>
                            <div class="form-group">
                                <label class="form-label">放行狀態:</label>
                                <div class="form-control">
                                    <asp:TextBox ID="txtApprovalStatus" runat="server" ReadOnly="true" CssClass="readonly-field"></asp:TextBox>
                                </div>
                            </div>
                            <div class="form-group">
                                <label class="form-label">核准人:</label>
                                <div class="form-control">
                                    <asp:TextBox ID="txtApprovedBy" runat="server" ReadOnly="true" CssClass="readonly-field"></asp:TextBox>
                                </div>
                            </div>
                        </div>
                    </div>

                    <!-- Tabs -->
                    <div class="tab-container">
                        <button type="button" class="tab-button active" id="btnTabExpense" onclick="switchTab('expense');">費用申請明細</button>
                        <button type="button" class="tab-button" id="btnTabMDR" onclick="switchTab('mdr');">MDR 發票明細</button>
                    </div>

                    <!-- Tab 1: Expense Lines -->
                    <div id="divContentExpense" class="tab-content active" runat="server" ClientIDMode="Static">
                        <div style="margin-bottom: 10px; display:flex; justify-content:space-between;">
                            <div>
                                <asp:Button ID="btnAddLine" runat="server" Text="+ 新增明細" OnClick="btnAddLine_Click" CssClass="btn btn-primary" />
                                <asp:Button ID="btnDeleteLine" runat="server" Text="🗑 刪除選取" OnClick="btnDeleteLine_Click" CssClass="btn btn-danger" OnClientClick="return confirm('確定刪除選中的明細行？');" />
                            </div>
                            <div>
                                <asp:FileUpload ID="fileUpload" runat="server" style="display:inline-block; width:200px;" />
                                <asp:Button ID="btnUpload" runat="server" Text="上傳附件" OnClick="btnUpload_Click" CssClass="btn btn-secondary btn-icon" />
                                <asp:Label ID="lblAttachment" runat="server" Text="" style="margin-left:5px;"></asp:Label>
                            </div>
                        </div>
                        
                        <div style="overflow-x:auto;">
                            <asp:GridView ID="gvExpenseDetail" runat="server" AutoGenerateColumns="False" CssClass="gridview"
                                         OnRowDataBound="gvExpenseDetail_RowDataBound" OnRowCommand="gvExpenseDetail_RowCommand">
                                <Columns>
                                    <asp:TemplateField HeaderText="選">
                                        <ItemTemplate>
                                            <asp:CheckBox ID="chkSelect" runat="server" />
                                        </ItemTemplate>
                                        <ItemStyle Width="30px" HorizontalAlign="Center" />
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="#">
                                        <ItemTemplate>
                                            <asp:Label ID="lblLineNum" runat="server" Text='<%# Container.DataItemIndex + 1 %>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="40px" HorizontalAlign="Center" />
                                    </asp:TemplateField>
                                    
                                    <asp:TemplateField HeaderText="費用類別">
                                        <ItemTemplate>
                                            <asp:DropDownList ID="ddlExpCategory" runat="server" Width="150px" AutoPostBack="true" OnSelectedIndexChanged="ddlExpCategory_SelectedIndexChanged"></asp:DropDownList>
                                        </ItemTemplate>
                                        <ItemStyle Width="160px" />
                                    </asp:TemplateField>
                                    
                                    <asp:TemplateField HeaderText="說明">
                                        <ItemTemplate>
                                            <asp:TextBox ID="txtDescription" runat="server" Width="200px" AutoPostBack="true" OnTextChanged="txtDescription_TextChanged"></asp:TextBox>
                                        </ItemTemplate>
                                        <ItemStyle Width="210px" />
                                    </asp:TemplateField>
                                    
                                    <asp:TemplateField HeaderText="會計科目">
                                        <ItemTemplate>
                                            <asp:TextBox ID="txtAcctCode" runat="server" Width="80px" ReadOnly="true" CssClass="readonly-field"></asp:TextBox>
                                        </ItemTemplate>
                                        <ItemStyle Width="90px" />
                                    </asp:TemplateField>
                                    
                                    <asp:TemplateField HeaderText="未稅金額">
                                        <ItemTemplate>
                                            <asp:TextBox ID="txtLineTotal" runat="server" Width="90px" style="text-align:right;" AutoPostBack="true" OnTextChanged="CalculateLineTotal"></asp:TextBox>
                                        </ItemTemplate>
                                        <ItemStyle Width="100px" />
                                    </asp:TemplateField>
                                    
                                    <asp:TemplateField HeaderText="稅別">
                                        <ItemTemplate>
                                            <asp:DropDownList ID="ddlVatGroup" runat="server" Width="80px" AutoPostBack="true" OnSelectedIndexChanged="CalculateLineTotal"></asp:DropDownList>
                                        </ItemTemplate>
                                        <ItemStyle Width="90px" />
                                    </asp:TemplateField>
                                    
                                    <asp:TemplateField HeaderText="稅額">
                                        <ItemTemplate>
                                            <asp:Label ID="lblVatSum" runat="server" Text="0" style="display:block; text-align:right;"></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="80px" />
                                    </asp:TemplateField>
                                    
                                    <asp:TemplateField HeaderText="含稅金額">
                                        <ItemTemplate>
                                            <asp:TextBox ID="txtPriceAfterVat" runat="server" Width="90px" style="text-align:right;" AutoPostBack="true" OnTextChanged="CalculatePriceAfterVat"></asp:TextBox>
                                        </ItemTemplate>
                                        <ItemStyle Width="100px" />
                                    </asp:TemplateField>
                                    
                                    <asp:TemplateField HeaderText="產品">
                                        <ItemTemplate>
                                            <asp:DropDownList ID="ddlCostingCode" runat="server" Width="100px"></asp:DropDownList>
                                        </ItemTemplate>
                                        <ItemStyle Width="110px" />
                                    </asp:TemplateField>
                                    
                                    <asp:TemplateField HeaderText="部門">
                                        <ItemTemplate>
                                            <asp:DropDownList ID="ddlCostingCode2" runat="server" Width="100px"></asp:DropDownList>
                                        </ItemTemplate>
                                        <ItemStyle Width="110px" />
                                    </asp:TemplateField>
                                </Columns>
                                <EmptyDataTemplate>
                                    <div style="text-align:center; padding:20px; color:gray;">請新增費用明細</div>
                                </EmptyDataTemplate>
                            </asp:GridView>
                        </div>
                    </div>

                    <!-- Tab 2: MDR Invoice Details -->
                    <div id="divContentMDR" class="tab-content" runat="server" ClientIDMode="Static">
                        <div style="margin-bottom: 10px;">
                            <asp:Button ID="btnAddMDRRow" runat="server" Text="+ 新增發票" OnClick="btnAddMDRRow_Click" CssClass="btn btn-primary" />
                            <asp:Button ID="btnDeleteMDRRow" runat="server" Text="🗑 刪除選取" OnClick="btnDeleteMDRRow_Click" CssClass="btn btn-danger" />
                        </div>
                        
                        <div style="overflow-x:auto;">
                            <asp:GridView ID="gvMDRDetail" runat="server" AutoGenerateColumns="False" CssClass="gridview"
                                         OnRowDataBound="gvMDRDetail_RowDataBound">
                                <Columns>
                                    <asp:TemplateField HeaderText="選">
                                        <ItemTemplate>
                                            <asp:CheckBox ID="chkSelectMDR" runat="server" />
                                        </ItemTemplate>
                                        <ItemStyle Width="30px" HorizontalAlign="Center" />
                                    </asp:TemplateField>
                                    
                                    <asp:TemplateField HeaderText="供應商代碼">
                                        <ItemTemplate>
                                            <asp:TextBox ID="txtLIFNR" runat="server" Text='<%# Bind("U_LIFNR") %>' Width="90px"></asp:TextBox>
                                        </ItemTemplate>
                                        <ItemStyle Width="100px" />
                                    </asp:TemplateField>

                                    <asp:TemplateField HeaderText="統一編號">
                                        <ItemTemplate>
                                            <asp:TextBox ID="txtSTCEG" runat="server" Text='<%# Bind("U_STCEG") %>' Width="90px"></asp:TextBox>
                                        </ItemTemplate>
                                        <ItemStyle Width="100px" />
                                    </asp:TemplateField>

                                    <asp:TemplateField HeaderText="發票號碼">
                                        <ItemTemplate>
                                            <asp:TextBox ID="txtXBLNR" runat="server" Text='<%# Bind("U_XBLNR") %>' Width="110px"></asp:TextBox>
                                        </ItemTemplate>
                                        <ItemStyle Width="120px" />
                                    </asp:TemplateField>

                                    <asp:TemplateField HeaderText="發票類型">
                                        <ItemTemplate>
                                            <asp:DropDownList ID="ddlZFORM_CODE" runat="server" SelectedValue='<%# Bind("U_ZFORM_CODE") %>' Width="150px">
                                                <asp:ListItem Value="21" Text="21-三聯式發票"></asp:ListItem>
                                                <asp:ListItem Value="22" Text="22-二聯式發票"></asp:ListItem>
                                                <asp:ListItem Value="25" Text="25-三聯式收銀機/電子發票"></asp:ListItem>
                                                <asp:ListItem Value="26" Text="26-三聯式/電子式/統一發票"></asp:ListItem>
                                                <asp:ListItem Value="27" Text="27-二聯式發票/普通收據"></asp:ListItem>
                                                <asp:ListItem Value="28" Text="28-載有稅額"></asp:ListItem>
                                            </asp:DropDownList>
                                        </ItemTemplate>
                                        <ItemStyle Width="160px" />
                                    </asp:TemplateField>

                                    <asp:TemplateField HeaderText="憑證日期">
                                        <ItemTemplate>
                                            <asp:TextBox ID="txtBLDAT" runat="server" Text='<%# Bind("U_BLDAT", "{0:yyyy-MM-dd}") %>' TextMode="Date" Width="120px"></asp:TextBox>
                                        </ItemTemplate>
                                        <ItemStyle Width="130px" />
                                    </asp:TemplateField>

                                    <asp:TemplateField HeaderText="營業稅日期">
                                        <ItemTemplate>
                                            <asp:TextBox ID="txtVATDATE" runat="server" Text='<%# Bind("U_VATDATE", "{0:yyyy-MM-dd}") %>' TextMode="Date" Width="120px"></asp:TextBox>
                                        </ItemTemplate>
                                        <ItemStyle Width="130px" />
                                    </asp:TemplateField>

                                    <asp:TemplateField HeaderText="未稅金額">
                                        <ItemTemplate>
                                            <asp:TextBox ID="txtHWBAS" runat="server" Text='<%# Bind("U_HWBAS", "{0:N2}") %>' Width="90px" style="text-align:right;" AutoPostBack="true" OnTextChanged="CalculateMDRTotal"></asp:TextBox>
                                        </ItemTemplate>
                                        <ItemStyle Width="100px" />
                                    </asp:TemplateField>

                                    <asp:TemplateField HeaderText="稅別">
                                        <ItemTemplate>
                                            <asp:DropDownList ID="ddlTAX_TYPE" runat="server" SelectedValue='<%# Bind("U_TAX_TYPE") %>' AutoPostBack="true" OnSelectedIndexChanged="CalculateMDRTotal" Width="80px">
                                                <asp:ListItem Value="1" Text="1-應稅"></asp:ListItem>
                                                <asp:ListItem Value="2" Text="2-零稅"></asp:ListItem>
                                                <asp:ListItem Value="3" Text="3-免稅"></asp:ListItem>
                                            </asp:DropDownList>
                                        </ItemTemplate>
                                        <ItemStyle Width="90px" />
                                    </asp:TemplateField>

                                    <asp:TemplateField HeaderText="稅額">
                                        <ItemTemplate>
                                            <asp:TextBox ID="txtHWSTE" runat="server" Text='<%# Bind("U_HWSTE", "{0:N2}") %>' Width="80px" style="text-align:right;" ReadOnly="true" CssClass="readonly-field"></asp:TextBox>
                                        </ItemTemplate>
                                        <ItemStyle Width="90px" />
                                    </asp:TemplateField>
                                </Columns>
                                <EmptyDataTemplate>
                                    <div style="text-align:center; padding:20px; color:gray;">請新增 MDR 發票明細</div>
                                </EmptyDataTemplate>
                            </asp:GridView>
                        </div>
                    </div>

                    <!-- Footer Section -->
                    <div style="margin-top:20px; border-top: 1px solid #ccc; padding-top:10px;">
                        <div class="row">
                            <div class="col-half">
                                <div class="form-group">
                                    <label class="form-label">採購人員:</label>
                                        <div class="form-control">
                                            <asp:DropDownList ID="ddlPurchaser" runat="server"></asp:DropDownList>
                                        </div>
                                    </div>
                                <div class="form-group">
                                    <label class="form-label">所有人:</label>
                                    <div class="form-control">
                                        <asp:TextBox ID="txtOwner" runat="server" CssClass="readonly-field" ReadOnly="true"></asp:TextBox>
                                    </div>
                                </div>
                                <div class="form-group">
                                    <label class="form-label">備註:</label>
                                    <div class="form-control">
                                        <asp:TextBox ID="txtRemarks" runat="server" TextMode="MultiLine" Height="50px"></asp:TextBox>
                                    </div>
                                </div>
                            </div>
                            <div class="col-half" style="text-align:right;">
                                <div class="form-group" style="justify-content: flex-end;">
                                    <label class="form-label" style="width:auto;">單據總額 (含稅):</label>
                                    <div style="width: 150px; margin-left:10px;">
                                        <asp:Label ID="lblDocTotalWithTax" runat="server" Text="0.00" Font-Bold="True" Font-Size="20px" ForeColor="Blue"></asp:Label>
                                    </div>
                                </div>
                                <div style="color:gray; font-size:12px;">
                                    未稅: <asp:Label ID="lblDocTotal" runat="server" Text="0.00"></asp:Label> | 
                                    稅額: <asp:Label ID="lblVatSum" runat="server" Text="0.00"></asp:Label>
                                </div>
                            </div>
                        </div>
                    </div>

                    <!-- Buttons -->
                    <div style="text-align:center; margin-top:30px;">
                        <asp:Button ID="btnSave" runat="server" Text="暫存 (Draft)" OnClick="btnSave_Click" CssClass="btn btn-primary" />
                        <asp:Button ID="btnSubmit" runat="server" Text="送出 (Submit)" OnClick="btnSubmit_Click" CssClass="btn btn-success" OnClientClick="return confirm('確定要送出審核嗎？');" />
                        <asp:Button ID="btnDelete" runat="server" Text="刪除 (Delete)" OnClick="btnDelete_Click" CssClass="btn btn-danger" OnClientClick="return confirmDelete();" />
                        <asp:Button ID="btnCancel" runat="server" Text="取消 (Cancel)" OnClick="btnCancel_Click" CssClass="btn btn-secondary" />
                        
                        <div style="margin-top:10px;">
                            <asp:Label ID="lblMessage" runat="server" Font-Bold="True"></asp:Label>
                        </div>
                    </div>

                    <!-- Approval Section (Hidden by default) -->
                    <asp:Panel ID="pnlApproval" runat="server" Visible="false" style="margin-top:20px; padding:15px; background-color:#fff3cd; border:1px solid #ffc107; border-radius:5px;">
                        <h3 style="margin-top:0;">審核意見</h3>
                        <asp:TextBox ID="txtApprovalComments" runat="server" TextMode="MultiLine" Height="60px"></asp:TextBox>
                        <div style="text-align:center; margin-top:10px;">
                            <asp:Button ID="btnApprove" runat="server" Text="核准" OnClick="btnApprove_Click" CssClass="btn btn-success" />
                            <asp:Button ID="btnReject" runat="server" Text="駁回" OnClick="btnReject_Click" CssClass="btn btn-danger" />
                        </div>
                    </asp:Panel>
                </div>

                <!-- Vendor Search Modal -->
                <asp:Button ID="btnDummy" runat="server" style="display:none" />
                <ajaxToolkit:ModalPopupExtender ID="mpeVendor" runat="server" TargetControlID="btnDummy"
                    PopupControlID="pnlVendorSearch" BackgroundCssClass="modalBackground" CancelControlID="btnCloseVendor" />
                <asp:Panel ID="pnlVendorSearch" runat="server" CssClass="modalPopup" style="display:none;">
                    <div class="modalHeader">
                        <span>供應商搜尋</span>
                        <asp:LinkButton ID="btnCloseVendor" runat="server" ForeColor="White" Font-Bold="true" style="text-decoration:none;">X</asp:LinkButton>
                    </div>
                    <div class="modalBody">
                        <div style="margin-bottom:10px;">
                            <div style="display:flex; align-items:center;">
                                <asp:TextBox ID="txtVendorSearchKeyword" runat="server" placeholder="輸入關鍵字..."></asp:TextBox>
                                <asp:Button ID="btnDoSearchVendor" runat="server" Text="搜尋" OnClick="btnDoSearchVendor_Click" CssClass="btn btn-primary" style="margin-left:5px;" />
                                <asp:HiddenField ID="hfSearchSource" runat="server" />
                            </div>
                            <div style="margin-top:5px;">
                                <asp:RadioButtonList ID="rblSearchMode" runat="server" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="Fuzzy" Selected="True">模糊搜尋</asp:ListItem>
                                    <asp:ListItem Value="Exact">完全比對</asp:ListItem>
                                </asp:RadioButtonList>
                            </div>
                        </div>
                        <asp:GridView ID="gvVendorSearch" runat="server" AutoGenerateColumns="False" Width="100%" CssClass="gridview"
                                     OnRowCommand="gvVendorSearch_RowCommand" AllowPaging="True" PageSize="10" OnPageIndexChanging="gvVendorSearch_PageIndexChanging">
                            <Columns>
                                <asp:TemplateField HeaderText="動作">
                                    <ItemTemplate>
                                        <asp:LinkButton ID="lbtnSelect" runat="server" CommandName="SelectVendor"
                                                       CommandArgument='<%# Eval("CardCode") + "|" + Eval("CardName") %>' CssClass="btn btn-success btn-icon">選取</asp:LinkButton>
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

            </ContentTemplate>
        </asp:UpdatePanel>
    </form>
</body>
</html>
