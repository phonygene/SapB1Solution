<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ExpenseClaimForm.aspx.vb" Inherits="MgmSP.ExpenseClaimForm" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>費用申請單</title>
    <style type="text/css">
        body {
            font-family: 'Microsoft JhengHei', Arial, sans-serif;
            font-size: 14px;
            margin: 0;
            padding: 0;
        }

        .container {
            width: 95%;
            margin: 20px auto;
        }

        .header {
            background-color: #0066CC;
            color: white;
            padding: 15px;
            border-radius: 5px 5px 0 0;
        }

        .toolbar {
            background-color: #F0F0F0;
            padding: 10px;
            border-bottom: 1px solid #CCC;
        }

        .toolbar button {
            margin-right: 10px;
            padding: 8px 15px;
            border: 1px solid #999;
            background-color: #FFFFFF;
            cursor: pointer;
            border-radius: 3px;
        }

        .toolbar button:hover {
            background-color: #E0E0E0;
        }

        .section {
            border: 1px solid #CCC;
            margin-top: 10px;
            border-radius: 5px;
        }

        .section-header {
            background-color: #E8E8E8;
            padding: 8px 15px;
            font-weight: bold;
            border-bottom: 1px solid #CCC;
        }

        .section-body {
            padding: 15px;
        }

        .form-row {
            margin-bottom: 10px;
            display: flex;
            align-items: center;
        }

        .form-label {
            width: 150px;
            font-weight: bold;
            text-align: right;
            margin-right: 10px;
        }

        .form-input {
            flex: 1;
            padding: 5px;
            border: 1px solid #CCC;
            border-radius: 3px;
        }

        .form-input[readonly] {
            background-color: #F0F0F0;
        }

        .form-input.required {
            background-color: #FFEEEE;
        }

        .form-input.searchable {
            background-color: #FFFFCC;
        }

        .grid-container {
            margin-top: 10px;
            overflow-x: auto;
        }

        .status-bar {
            background-color: #F8F8F8;
            border-top: 1px solid #CCC;
            padding: 8px 15px;
            text-align: right;
            font-size: 12px;
            color: #666;
        }

        .error-message {
            color: red;
            font-weight: bold;
            margin-top: 10px;
        }

        .success-message {
            color: green;
            font-weight: bold;
            margin-top: 10px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server"></asp:ScriptManager>

        <div class="container">
            <!-- 標題列 -->
            <div class="header">
                <h2 style="margin: 0;">費用申請單（Expense Claim）</h2>
                <span id="lblMode" runat="server">模式：新增（Create）</span>
            </div>

            <!-- 工具列 -->
            <div class="toolbar">
                <asp:Button ID="btnCreate" runat="server" Text="新增 (Create)" OnClick="btnCreate_Click" />
                <asp:Button ID="btnSearch" runat="server" Text="搜尋 (Search)" OnClick="btnSearch_Click" />
                <asp:Button ID="btnUpdate" runat="server" Text="更新 (Update)" OnClick="btnUpdate_Click" Enabled="false" />
                <asp:Button ID="btnSave" runat="server" Text="儲存 (Save)" OnClick="btnSave_Click" />
                <asp:Button ID="btnCancel" runat="server" Text="取消 (Cancel)" OnClick="btnCancel_Click" />
            </div>

            <!-- 訊息區 -->
            <asp:Label ID="lblMessage" runat="server" CssClass="error-message" Visible="false"></asp:Label>

            <!-- 表頭區 -->
            <div class="section">
                <div class="section-header">表頭資訊</div>
                <div class="section-body">
                    <div class="form-row">
                        <span class="form-label">ID:</span>
                        <asp:TextBox ID="txtID" runat="server" CssClass="form-input" ReadOnly="true"></asp:TextBox>
                    </div>
                    <div class="form-row">
                        <span class="form-label">AP 發票 DocEntry:</span>
                        <asp:TextBox ID="txtDocEntry" runat="server" CssClass="form-input" ReadOnly="true"></asp:TextBox>
                    </div>
                    <div class="form-row">
                        <span class="form-label">AP 發票單號:</span>
                        <asp:TextBox ID="txtDocNum" runat="server" CssClass="form-input" ReadOnly="true"></asp:TextBox>
                    </div>
                    <div class="form-row">
                        <span class="form-label">單據總金額:</span>
                        <asp:TextBox ID="txtDocTotal" runat="server" CssClass="form-input" ReadOnly="true" Text="0.00"></asp:TextBox>
                    </div>
                    <div class="form-row">
                        <span class="form-label">發票稅額總金額:</span>
                        <asp:TextBox ID="txtVatSum" runat="server" CssClass="form-input" ReadOnly="true" Text="0.00"></asp:TextBox>
                    </div>
                    <div class="form-row">
                        <span class="form-label">建立日期:</span>
                        <asp:TextBox ID="txtCreateDate" runat="server" CssClass="form-input" ReadOnly="true"></asp:TextBox>
                    </div>
                    <div class="form-row">
                        <span class="form-label">建立人:</span>
                        <asp:TextBox ID="txtCreateBy" runat="server" CssClass="form-input" ReadOnly="true"></asp:TextBox>
                    </div>
                </div>
            </div>

            <!-- 發票明細區 -->
            <div class="section">
                <div class="section-header">
                    發票明細
                    <asp:Button ID="btnAddRow" runat="server" Text="新增列" OnClick="btnAddRow_Click" style="float: right; margin-top: -3px;" />
                </div>
                <div class="section-body">
                    <div class="grid-container">
                        <asp:GridView ID="gvInvoiceDetail" runat="server"
                            AutoGenerateColumns="False"
                            CellPadding="4"
                            ForeColor="#333333"
                            GridLines="Both"
                            Width="100%"
                            OnRowDataBound="gvInvoiceDetail_RowDataBound"
                            OnRowCommand="gvInvoiceDetail_RowCommand"
                            DataKeyNames="LineId">
                            <AlternatingRowStyle BackColor="White" />
                            <Columns>
                                <asp:BoundField DataField="LineId" HeaderText="列號" ReadOnly="True" />

                                <asp:TemplateField HeaderText="供應商代碼*">
                                    <ItemTemplate>
                                        <asp:TextBox ID="txtLIFNR" runat="server" Text='<%# Bind("U_LIFNR") %>' Width="100px"></asp:TextBox>
                                    </ItemTemplate>
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="統一編號*">
                                    <ItemTemplate>
                                        <asp:TextBox ID="txtSTCEG" runat="server" Text='<%# Bind("U_STCEG") %>' Width="100px"></asp:TextBox>
                                    </ItemTemplate>
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="發票號碼*">
                                    <ItemTemplate>
                                        <asp:TextBox ID="txtXBLNR" runat="server" Text='<%# Bind("U_XBLNR") %>' Width="120px"></asp:TextBox>
                                    </ItemTemplate>
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="發票類型*">
                                    <ItemTemplate>
                                        <asp:DropDownList ID="ddlZFORM_CODE" runat="server" SelectedValue='<%# Bind("U_ZFORM_CODE") %>'>
                                            <asp:ListItem Value="21" Text="21-三聯式發票"></asp:ListItem>
                                            <asp:ListItem Value="22" Text="22-二聯式發票"></asp:ListItem>
                                            <asp:ListItem Value="25" Text="25-三聯式收銀機/電子發票"></asp:ListItem>
                                            <asp:ListItem Value="26" Text="26-三聯式/電子式/統一發票"></asp:ListItem>
                                            <asp:ListItem Value="27" Text="27-二聯式發票/普通收據"></asp:ListItem>
                                            <asp:ListItem Value="28" Text="28-載有稅額"></asp:ListItem>
                                        </asp:DropDownList>
                                    </ItemTemplate>
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="憑證日期*">
                                    <ItemTemplate>
                                        <asp:TextBox ID="txtBLDAT" runat="server" Text='<%# Bind("U_BLDAT", "{0:yyyy-MM-dd}") %>' Width="100px"></asp:TextBox>
                                    </ItemTemplate>
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="營業稅日期*">
                                    <ItemTemplate>
                                        <asp:TextBox ID="txtVATDATE" runat="server" Text='<%# Bind("U_VATDATE", "{0:yyyy-MM-dd}") %>' Width="100px"></asp:TextBox>
                                    </ItemTemplate>
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="未稅金額*">
                                    <ItemTemplate>
                                        <asp:TextBox ID="txtHWBAS" runat="server" Text='<%# Bind("U_HWBAS", "{0:N2}") %>' Width="100px" AutoPostBack="true" OnTextChanged="txtHWBAS_TextChanged"></asp:TextBox>
                                    </ItemTemplate>
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="稅額*">
                                    <ItemTemplate>
                                        <asp:TextBox ID="txtHWSTE" runat="server" Text='<%# Bind("U_HWSTE", "{0:N2}") %>' Width="100px" ReadOnly="true"></asp:TextBox>
                                    </ItemTemplate>
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="稅別*">
                                    <ItemTemplate>
                                        <asp:DropDownList ID="ddlTAX_TYPE" runat="server" SelectedValue='<%# Bind("U_TAX_TYPE") %>' AutoPostBack="true" OnSelectedIndexChanged="ddlTAX_TYPE_SelectedIndexChanged">
                                            <asp:ListItem Value="1" Text="1-應稅"></asp:ListItem>
                                            <asp:ListItem Value="2" Text="2-零稅"></asp:ListItem>
                                            <asp:ListItem Value="3" Text="3-免稅"></asp:ListItem>
                                        </asp:DropDownList>
                                    </ItemTemplate>
                                </asp:TemplateField>

                                <asp:ButtonField CommandName="Delete" Text="刪除" ButtonType="Button" />
                            </Columns>
                            <EditRowStyle BackColor="#2461BF" />
                            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                            <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                            <RowStyle BackColor="#EFF3FB" />
                            <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                            <SortedAscendingCellStyle BackColor="#F5F7FB" />
                            <SortedAscendingHeaderStyle BackColor="#6D95E1" />
                            <SortedDescendingCellStyle BackColor="#E9EBEF" />
                            <SortedDescendingHeaderStyle BackColor="#4870BE" />
                        </asp:GridView>
                    </div>
                </div>
            </div>

            <!-- 狀態列 -->
            <div class="status-bar">
                <span id="lblStatus" runat="server">就緒</span> |
                使用者: <asp:Label ID="lblUser" runat="server"></asp:Label> |
                時間: <asp:Label ID="lblTime" runat="server"></asp:Label>
            </div>
        </div>
    </form>
</body>
</html>
