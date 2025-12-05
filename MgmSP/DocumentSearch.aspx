<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="DocumentSearch.aspx.vb" Inherits="MgmSP.DocumentSearch"
    MaintainScrollPositionOnPostback="true" MasterPageFile="~/MySite1.Master" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <title>單據查詢</title>
    <style type="text/css">
        body {
            font-family: "Microsoft JhengHei", Arial, sans-serif;
            font-size: 14px;
            background-color: #f5f5f5;
        }

        .form-container {
            max-width: 1400px;
            margin: 20px auto;
            padding: 20px;
            background-color: white;
            box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
            border-radius: 5px;
        }

        .section-header {
            background-color: #17a2b8;
            color: white;
            padding: 10px 15px;
            margin: 10px 0 15px 0;
            font-weight: bold;
            border-radius: 4px;
            font-size: 16px;
        }

        .filter-row {
            display: flex;
            flex-wrap: wrap;
            gap: 15px;
            margin-bottom: 10px;
        }

        .filter-group {
            display: flex;
            align-items: center;
        }

        .filter-label {
            font-weight: bold;
            margin-right: 5px;
            white-space: nowrap;
            color: #333;
        }

        .filter-control input[type="text"],
        .filter-control select {
            padding: 5px 8px;
            border: 1px solid #ccc;
            border-radius: 4px;
            font-size: 13px;
        }

        .filter-control input[type="date"] {
            padding: 4px 6px;
        }

        .range-sep {
            margin: 0 3px;
            color: #666;
        }

        .btn {
            padding: 8px 20px;
            border-radius: 4px;
            border: none;
            cursor: pointer;
            font-size: 14px;
            font-weight: bold;
            margin-right: 5px;
        }

        .btn-primary {
            background-color: #007bff;
            color: white;
        }

        .btn-secondary {
            background-color: #6c757d;
            color: white;
        }

        .btn:hover {
            opacity: 0.9;
        }

        .readonly-field {
            background-color: #e9ecef;
        }

        .gridview {
            border-collapse: collapse;
            width: 100%;
            margin-top: 15px;
            font-size: 13px;
        }

        .gridview th {
            background-color: #17a2b8;
            color: white;
            padding: 10px;
            border: 1px solid #ddd;
            text-align: center;
            white-space: nowrap;
        }

        .gridview td {
            padding: 8px;
            border: 1px solid #ddd;
            vertical-align: middle;
        }

        .gridview tr:nth-child(even) {
            background-color: #f8f9fa;
        }

        .gridview tr:hover {
            background-color: #e9ecef;
        }

        .link-jid {
            color: #007bff;
            text-decoration: underline;
            cursor: pointer;
            font-weight: bold;
        }

        .link-jid:hover {
            color: #0056b3;
        }

        .badge {
            padding: 3px 8px;
            border-radius: 10px;
            color: white;
            font-size: 11px;
            font-weight: bold;
        }

        .status-P {
            background-color: #6c757d;
        }

        .status-W {
            background-color: #ffc107;
            color: black;
        }

        .status-A {
            background-color: #28a745;
        }

        .status-R {
            background-color: #dc3545;
        }

        .remarks-cell {
            max-width: 200px;
            cursor: help;
        }

        .pager {
            margin-top: 15px;
            text-align: center;
        }

        .pager a,
        .pager span {
            padding: 5px 10px;
            margin: 0 2px;
            border: 1px solid #ddd;
            text-decoration: none;
        }

        .pager a:hover {
            background-color: #e9ecef;
        }

        .pager span {
            background-color: #17a2b8;
            color: white;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server"></asp:ScriptManager>

    <div class="form-container">
        <h2 style="margin:0 0 20px 0; color:#17a2b8;">單據查詢</h2>

        <div class="section-header">篩選條件</div>

        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
            <ContentTemplate>
                <!-- Row 1 -->
                <div class="filter-row">
                    <div class="filter-group">
                        <span class="filter-label">單據類型:</span>
                        <div class="filter-control">
                            <asp:DropDownList ID="ddlDocType" runat="server" Width="120px">
                                <asp:ListItem Value="ExpenseClaim" Text="費用申請單"></asp:ListItem>
                            </asp:DropDownList>
                        </div>
                    </div>
                    <div class="filter-group">
                        <span class="filter-label">使用者代碼:</span>
                        <div class="filter-control">
                            <asp:TextBox ID="txtUserCode" runat="server" Width="80px"></asp:TextBox>
                        </div>
                    </div>
                    <div class="filter-group">
                        <span class="filter-label">使用者名稱:</span>
                        <div class="filter-control">
                            <asp:TextBox ID="txtUserName" runat="server" Width="100px"></asp:TextBox>
                        </div>
                    </div>
                    <div class="filter-group">
                        <span class="filter-label">jID:</span>
                        <div class="filter-control">
                            <asp:TextBox ID="txtJID" runat="server" Width="80px"></asp:TextBox>
                        </div>
                    </div>
                </div>

                <!-- Row 2 -->
                <div class="filter-row">
                    <div class="filter-group">
                        <span class="filter-label">AP單號:</span>
                        <div class="filter-control">
                            <asp:TextBox ID="txtAPNumFrom" runat="server" Width="80px" placeholder="起"></asp:TextBox>
                            <span class="range-sep">~</span>
                            <asp:TextBox ID="txtAPNumTo" runat="server" Width="80px" placeholder="迄"></asp:TextBox>
                        </div>
                    </div>
                    <div class="filter-group">
                        <span class="filter-label">簽核系統PID:</span>
                        <div class="filter-control">
                            <asp:TextBox ID="txtPIDFrom" runat="server" Width="80px" placeholder="起"></asp:TextBox>
                            <span class="range-sep">~</span>
                            <asp:TextBox ID="txtPIDTo" runat="server" Width="80px" placeholder="迄"></asp:TextBox>
                        </div>
                    </div>
                    <div class="filter-group">
                        <span class="filter-label">文件狀態:</span>
                        <div class="filter-control">
                            <asp:DropDownList ID="ddlDocStatus" runat="server" Width="100px">
                                <asp:ListItem Value="" Text="全部"></asp:ListItem>
                                <asp:ListItem Value="P" Text="草稿"></asp:ListItem>
                                <asp:ListItem Value="W" Text="待審核"></asp:ListItem>
                                <asp:ListItem Value="A" Text="已核准"></asp:ListItem>
                                <asp:ListItem Value="R" Text="已退回"></asp:ListItem>
                            </asp:DropDownList>
                        </div>
                    </div>
                    <div class="filter-group">
                        <span class="filter-label">放行狀態:</span>
                        <div class="filter-control">
                            <asp:DropDownList ID="ddlApprovalStatus" runat="server" Width="100px">
                                <asp:ListItem Value="" Text="全部"></asp:ListItem>
                                <asp:ListItem Value="Y" Text="已放行"></asp:ListItem>
                                <asp:ListItem Value="N" Text="未放行"></asp:ListItem>
                            </asp:DropDownList>
                        </div>
                    </div>
                </div>

                <!-- Row 3 -->
                <div class="filter-row">
                    <div class="filter-group">
                        <span class="filter-label">文件日期:</span>
                        <div class="filter-control">
                            <asp:TextBox ID="txtDocDateFrom" runat="server" TextMode="Date" Width="130px"></asp:TextBox>
                            <span class="range-sep">~</span>
                            <asp:TextBox ID="txtDocDateTo" runat="server" TextMode="Date" Width="130px"></asp:TextBox>
                        </div>
                    </div>
                    <div class="filter-group">
                        <span class="filter-label">到期日:</span>
                        <div class="filter-control">
                            <asp:TextBox ID="txtDueDateFrom" runat="server" TextMode="Date" Width="130px"></asp:TextBox>
                            <span class="range-sep">~</span>
                            <asp:TextBox ID="txtDueDateTo" runat="server" TextMode="Date" Width="130px"></asp:TextBox>
                        </div>
                    </div>
                    <div class="filter-group">
                        <span class="filter-label">過帳日期:</span>
                        <div class="filter-control">
                            <asp:TextBox ID="txtTaxDateFrom" runat="server" TextMode="Date" Width="130px"></asp:TextBox>
                            <span class="range-sep">~</span>
                            <asp:TextBox ID="txtTaxDateTo" runat="server" TextMode="Date" Width="130px"></asp:TextBox>
                        </div>
                    </div>
                </div>

                <!-- Row 4 -->
                <div class="filter-row">
                    <div class="filter-group">
                        <span class="filter-label">供應商代碼:</span>
                        <div class="filter-control">
                            <asp:TextBox ID="txtCardCodeFrom" runat="server" Width="80px" placeholder="起"></asp:TextBox>
                            <span class="range-sep">~</span>
                            <asp:TextBox ID="txtCardCodeTo" runat="server" Width="80px" placeholder="迄"></asp:TextBox>
                        </div>
                    </div>
                    <div class="filter-group">
                        <span class="filter-label">供應商名稱:</span>
                        <div class="filter-control">
                            <asp:TextBox ID="txtCardName" runat="server" Width="120px"></asp:TextBox>
                            <asp:RadioButtonList ID="rblCardNameMode" runat="server"
                                RepeatDirection="Horizontal" style="display:inline-block; margin-left:5px;">
                                <asp:ListItem Value="StartsWith" Text="開頭" Selected="True"></asp:ListItem>
                                <asp:ListItem Value="Contains" Text="模糊"></asp:ListItem>
                            </asp:RadioButtonList>
                        </div>
                    </div>
                    <div class="filter-group">
                        <span class="filter-label">備註:</span>
                        <div class="filter-control">
                            <asp:TextBox ID="txtComments" runat="server" Width="120px"></asp:TextBox>
                            <asp:RadioButtonList ID="rblCommentsMode" runat="server"
                                RepeatDirection="Horizontal" style="display:inline-block; margin-left:5px;">
                                <asp:ListItem Value="StartsWith" Text="開頭" Selected="True"></asp:ListItem>
                                <asp:ListItem Value="Contains" Text="模糊"></asp:ListItem>
                            </asp:RadioButtonList>
                        </div>
                    </div>
                </div>

                <!-- Row 5: Sorting & Paging -->
                <div class="filter-row" style="margin-top:15px; padding-top:15px; border-top:1px solid #ddd;">
                    <div class="filter-group">
                        <span class="filter-label">排序依據:</span>
                        <div class="filter-control">
                            <asp:DropDownList ID="ddlSortBy" runat="server" Width="120px">
                                <asp:ListItem Value="jID" Text="jID" Selected="True"></asp:ListItem>
                                <asp:ListItem Value="DocDate" Text="文件日期"></asp:ListItem>
                                <asp:ListItem Value="DocDueDate" Text="到期日"></asp:ListItem>
                                <asp:ListItem Value="CardCode" Text="供應商代碼"></asp:ListItem>
                                <asp:ListItem Value="CardName" Text="供應商名稱"></asp:ListItem>
                                <asp:ListItem Value="CreateDate" Text="建立日期"></asp:ListItem>
                            </asp:DropDownList>
                        </div>
                    </div>
                    <div class="filter-group">
                        <asp:RadioButtonList ID="rblSortOrder" runat="server" RepeatDirection="Horizontal">
                            <asp:ListItem Value="DESC" Text="倒序" Selected="True"></asp:ListItem>
                            <asp:ListItem Value="ASC" Text="正序"></asp:ListItem>
                        </asp:RadioButtonList>
                    </div>
                    <div class="filter-group">
                        <span class="filter-label">每頁筆數:</span>
                        <div class="filter-control">
                            <asp:DropDownList ID="ddlPageSize" runat="server" Width="70px">
                                <asp:ListItem Value="10" Text="10"></asp:ListItem>
                                <asp:ListItem Value="20" Text="20" Selected="True"></asp:ListItem>
                                <asp:ListItem Value="50" Text="50"></asp:ListItem>
                                <asp:ListItem Value="100" Text="100"></asp:ListItem>
                            </asp:DropDownList>
                        </div>
                    </div>
                    <div class="filter-group" style="margin-left:auto;">
                        <asp:Button ID="btnSearch" runat="server" Text="查詢" CssClass="btn btn-primary" OnClick="btnSearch_Click" />
                        <asp:Button ID="btnClear" runat="server" Text="清除條件" CssClass="btn btn-secondary" OnClick="btnClear_Click" />
                    </div>
                </div>

                <!-- Message -->
                <asp:Label ID="lblMessage" runat="server" Font-Bold="True"></asp:Label>

                <!-- Results Section -->
                <div class="section-header" style="margin-top:25px;">查詢結果</div>
                <asp:Label ID="lblResultCount" runat="server" style="color:#666; font-size:13px;"></asp:Label>

                <asp:GridView ID="gvResults" runat="server" AutoGenerateColumns="False" CssClass="gridview"
                    AllowPaging="True" OnPageIndexChanging="gvResults_PageIndexChanging"
                    OnRowDataBound="gvResults_RowDataBound">
                    <Columns>
                        <asp:TemplateField HeaderText="jID">
                            <ItemTemplate>
                                <asp:HyperLink ID="hlJID" runat="server" CssClass="link-jid"
                                    NavigateUrl='<%# "ExpenseClaimForm.aspx?DocEntry=" & Eval("jID") %>'
                                    Text='<%# Eval("jID") %>'></asp:HyperLink>
                            </ItemTemplate>
                            <ItemStyle HorizontalAlign="Center" Width="60px" />
                        </asp:TemplateField>
                        <asp:BoundField DataField="CardName" HeaderText="供應商名稱" />
                        <asp:BoundField DataField="InvNum" HeaderText="AP單號" />
                        <asp:BoundField DataField="U_PID" HeaderText="簽核系統PID" />
                        <asp:TemplateField HeaderText="文件狀態">
                            <ItemTemplate>
                                <span class='<%# "badge status-" & Eval("ApprovalStatus") %>'>
                                    <%# GetStatusText(Eval("ApprovalStatus").ToString()) %>
                                </span>
                            </ItemTemplate>
                            <ItemStyle HorizontalAlign="Center" Width="80px" />
                        </asp:TemplateField>
                        <asp:BoundField DataField="DocDate" HeaderText="文件日期" DataFormatString="{0:yyyy-MM-dd}" />
                        <asp:TemplateField HeaderText="放行狀態">
                            <ItemTemplate>
                                <%# If(Eval("IsApproved") IsNot Nothing AndAlso Eval("IsApproved").ToString() = "Y", "已放行", "未放行") %>
                            </ItemTemplate>
                            <ItemStyle HorizontalAlign="Center" Width="70px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="備註">
                            <ItemTemplate>
                                <div class="remarks-cell" title='<%# Eval("Comments") %>'>
                                    <%# TruncateRemarks(Eval("Comments").ToString(), 20) %>
                                </div>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="CreateBy" HeaderText="建立者" />
                    </Columns>
                    <PagerStyle CssClass="pager" />
                    <EmptyDataTemplate>
                        <div style="text-align:center; padding:30px; color:gray;">
                            請輸入篩選條件後按「查詢」
                        </div>
                    </EmptyDataTemplate>
                </asp:GridView>

            </ContentTemplate>
        </asp:UpdatePanel>
    </div>

    <!-- Hidden controls for modal -->
    <asp:Button ID="btnDummyRemarks" runat="server" style="display:none" />
    <ajaxToolkit:ModalPopupExtender ID="mpeRemarks" runat="server" TargetControlID="btnDummyRemarks"
        PopupControlID="pnlRemarks" BackgroundCssClass="modalBackground" CancelControlID="btnCloseRemarks" />
    <asp:Panel ID="pnlRemarks" runat="server" CssClass="modalPopup" style="display:none;">
        <div class="modalFooter">
            <asp:Button ID="btnCloseRemarks" runat="server" Text="關閉" CssClass="btn btn-secondary" />
        </div>
    </asp:Panel>
</asp:Content>
