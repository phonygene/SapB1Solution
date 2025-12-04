<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ExpenseClaimList.aspx.vb" Inherits="MgmSP.ExpenseClaimList" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>費用申請單查詢</title>
    <style type="text/css">
        body { font-family: "Microsoft JhengHei", Arial, sans-serif; font-size: 14px; background-color: #f5f5f5; }
        .container { max-width: 1400px; margin: 20px auto; padding: 20px; background-color: white; box-shadow: 0 0 10px rgba(0,0,0,0.1); border-radius: 5px; }
        
        .section-header { 
            background-color: #007bff; color: white; padding: 10px 15px;
            margin: 0 0 15px 0; font-weight: bold; border-radius: 4px; font-size: 16px;
        }

        .search-panel { border: 1px solid #ddd; padding: 15px; border-radius: 5px; background-color: #f9f9f9; margin-bottom: 20px; }
        
        .row { display: flex; flex-wrap: wrap; margin-bottom: 10px; align-items: center; }
        .col { flex: 1; padding: 0 10px; min-width: 200px; }
        .col-btn { text-align: right; padding: 0 10px; }
        
        .form-group { display: flex; align-items: center; margin-bottom: 5px; }
        .form-label { width: 100px; text-align: right; padding-right: 10px; font-weight: bold; color: #555; }
        .form-control { flex: 1; padding: 5px; border: 1px solid #ccc; border-radius: 3px; }
        
        input[type="text"], input[type="date"], select { width: 100%; box-sizing: border-box; padding: 6px; }
        
        .btn { padding: 8px 15px; border-radius: 4px; border: none; cursor: pointer; font-size: 14px; font-weight: bold; margin-left: 5px; color: white; }
        .btn-primary { background-color: #007bff; }
        .btn-success { background-color: #28a745; }
        .btn-secondary { background-color: #6c757d; }
        .btn:hover { opacity: 0.9; }

        .gridview { border-collapse: collapse; width: 100%; font-size: 13px; margin-top: 10px; }
        .gridview th { background-color: #007bff; color: white; padding: 10px; text-align: center; border: 1px solid #ddd; }
        .gridview td { padding: 8px; border: 1px solid #ddd; vertical-align: middle; color: #333; }
        .gridview tr:nth-child(even) { background-color: #f2f2f2; }
        .gridview tr:hover { background-color: #e9ecef; }
        
        .badge { padding: 4px 8px; border-radius: 10px; color: white; font-size: 12px; font-weight: bold; display: inline-block; }
        .status-P { background-color: #6c757d; } /* 草稿 */
        .status-W { background-color: #ffc107; color: black; } /* 待審 */
        .status-A { background-color: #28a745; } /* 核准 */
        .status-R { background-color: #dc3545; } /* 駁回 */
        
        .link-btn { color: #007bff; text-decoration: none; font-weight: bold; cursor: pointer; }
        .link-btn:hover { text-decoration: underline; }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <div class="container">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:20px;">
                <h2 style="margin:0; color:#333;">費用申請單查詢</h2>
                <asp:Button ID="btnCreateNew" runat="server" Text="+ 新增申請單" CssClass="btn btn-success" PostBackUrl="~/MgmSP/ExpenseClaimForm.aspx" />
            </div>

            <div class="search-panel">
                <div class="section-header">查詢條件</div>
                
                <div class="row">
                    <div class="col">
                        <div class="form-group">
                            <label class="form-label">jID:</label>
                            <asp:TextBox ID="txtJID" runat="server" CssClass="form-control" placeholder="平台單號"></asp:TextBox>
                        </div>
                    </div>
                    <div class="col">
                        <div class="form-group">
                            <label class="form-label">AP 單號:</label>
                            <asp:TextBox ID="txtDocEntry" runat="server" CssClass="form-control" placeholder="SAP DocEntry"></asp:TextBox>
                        </div>
                    </div>
                    <div class="col">
                        <div class="form-group">
                            <label class="form-label">簽核 PID:</label>
                            <asp:TextBox ID="txtUPID" runat="server" CssClass="form-control"></asp:TextBox>
                        </div>
                    </div>
                </div>

                <div class="row">
                    <div class="col">
                        <div class="form-group">
                            <label class="form-label">供應商:</label>
                            <asp:TextBox ID="txtVendor" runat="server" CssClass="form-control" placeholder="代碼或名稱"></asp:TextBox>
                        </div>
                    </div>
                    <div class="col">
                        <div class="form-group">
                            <label class="form-label">文件狀態:</label>
                            <asp:DropDownList ID="ddlStatus" runat="server" CssClass="form-control">
                                <asp:ListItem Value="" Text="全部"></asp:ListItem>
                                <asp:ListItem Value="P" Text="草稿 (Draft)"></asp:ListItem>
                                <asp:ListItem Value="W" Text="待審核 (Pending)"></asp:ListItem>
                                <asp:ListItem Value="A" Text="已核准 (Approved)"></asp:ListItem>
                                <asp:ListItem Value="R" Text="已駁回 (Rejected)"></asp:ListItem>
                            </asp:DropDownList>
                        </div>
                    </div>
                    <div class="col">
                        <div class="form-group">
                            <label class="form-label">備註:</label>
                            <asp:TextBox ID="txtRemarks" runat="server" CssClass="form-control"></asp:TextBox>
                        </div>
                    </div>
                </div>

                <div class="row">
                    <div class="col" style="flex: 2;">
                        <div class="form-group">
                            <label class="form-label">日期區間:</label>
                            <div style="display:flex; width:100%;">
                                <asp:DropDownList ID="ddlDateType" runat="server" style="width:120px; margin-right:5px;">
                                    <asp:ListItem Value="DocDate" Text="文件日期"></asp:ListItem>
                                    <asp:ListItem Value="DocDueDate" Text="到期日"></asp:ListItem>
                                    <asp:ListItem Value="TaxDate" Text="過帳日期"></asp:ListItem>
                                </asp:DropDownList>
                                <asp:TextBox ID="txtDateFrom" runat="server" TextMode="Date" style="width:140px; margin-right:5px;"></asp:TextBox>
                                <span style="padding:5px;">~</span>
                                <asp:TextBox ID="txtDateTo" runat="server" TextMode="Date" style="width:140px; margin-left:5px;"></asp:TextBox>
                            </div>
                        </div>
                    </div>
                    <div class="col col-btn">
                        <asp:Button ID="btnSearch" runat="server" Text="查詢" CssClass="btn btn-primary" OnClick="btnSearch_Click" />
                        <asp:Button ID="btnClear" runat="server" Text="清除" CssClass="btn btn-secondary" OnClick="btnClear_Click" />
                    </div>
                </div>
            </div>

            <asp:GridView ID="gvList" runat="server" AutoGenerateColumns="False" CssClass="gridview" 
                          AllowPaging="True" PageSize="20" OnPageIndexChanging="gvList_PageIndexChanging" OnRowDataBound="gvList_RowDataBound">
                <Columns>
                    <asp:TemplateField HeaderText="jID">
                        <ItemTemplate>
                            <asp:HyperLink ID="hlJID" runat="server" CssClass="link-btn"
                                NavigateUrl='<%# "ExpenseClaimForm.aspx?DocEntry=" + Eval("DocEntry").ToString() %>'
                                Text='<%# Eval("jID") %>'></asp:HyperLink>
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Center" Width="80px" />
                    </asp:TemplateField>
                    
                    <asp:BoundField DataField="DocEntry" HeaderText="AP 單號" ItemStyle-HorizontalAlign="Center" />
                    
                    <asp:TemplateField HeaderText="狀態">
                        <ItemTemplate>
                            <asp:Label ID="lblStatus" runat="server" Text='<%# Eval("ApprovalStatus") %>'></asp:Label>
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Center" Width="100px" />
                    </asp:TemplateField>

                    <asp:BoundField DataField="CardName" HeaderText="供應商名稱" />
                    <asp:BoundField DataField="DocDate" HeaderText="文件日期" DataFormatString="{0:yyyy-MM-dd}" ItemStyle-HorizontalAlign="Center" />
                    <asp:BoundField DataField="DocTotal" HeaderText="總金額" DataFormatString="{0:N2}" ItemStyle-HorizontalAlign="Right" />
                    <asp:BoundField DataField="Comments" HeaderText="備註" />
                    <asp:BoundField DataField="U_PID" HeaderText="簽核 PID" ItemStyle-HorizontalAlign="Center" />
                    <asp:BoundField DataField="CreateBy" HeaderText="建立者" ItemStyle-HorizontalAlign="Center" />
                </Columns>
                <EmptyDataTemplate>
                    <div style="text-align:center; padding:20px;">查無資料</div>
                </EmptyDataTemplate>
                <PagerStyle HorizontalAlign="Center" CssClass="gridview" BackColor="#e9ecef" />
            </asp:GridView>
        </div>
    </form>
</body>
</html>