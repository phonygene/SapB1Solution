<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="cncitems.aspx.vb" Inherits="MgmSP.cncitems" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:Table ID="HyperMenuT" runat="server">
            </asp:Table>
            <asp:Table ID="InfoT" runat="server" GridLines="Both" Width="100%" BackColor="#3399FF">
            </asp:Table>
            <asp:Table ID="CncAddItemT" runat="server" GridLines="Both" Width="100%" BackColor="#99FF66">
            </asp:Table>
            <asp:GridView ID="gv1" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" Width="100%">
                <AlternatingRowStyle BackColor="#FFFFCC" />
                <Columns>
                    <asp:BoundField DataField="no" HeaderText="項次">
                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                    <ItemStyle Font-Size="Smaller" HorizontalAlign="Center" VerticalAlign="Middle" />
                    </asp:BoundField>
                    <asp:BoundField DataField="unitno" HeaderText="刻字">
                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                    <ItemStyle Font-Size="Smaller" HorizontalAlign="Center" VerticalAlign="Middle" />
                    </asp:BoundField>
                    <asp:BoundField DataField="itemcode" HeaderText="料號">
                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                    <ItemStyle Font-Size="Smaller" HorizontalAlign="Left" VerticalAlign="Middle" />
                    </asp:BoundField>
                    <asp:BoundField DataField="itemname" HeaderText="說明">
                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                    <ItemStyle Font-Size="Smaller" HorizontalAlign="Left" VerticalAlign="Middle" />
                    </asp:BoundField>
                    <asp:BoundField DataField="rawspec" HeaderText="原料規格">
                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                    <ItemStyle Font-Size="Smaller" HorizontalAlign="Left" VerticalAlign="Middle" />
                    </asp:BoundField>
                    <asp:BoundField DataField="stat" HeaderText="狀態">
                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                    <ItemStyle Font-Size="Smaller" HorizontalAlign="Center" VerticalAlign="Middle" />
                    </asp:BoundField>
                    <asp:BoundField DataField="quantity" HeaderText="加工數量" DataFormatString="{0:F}">
                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                    <ItemStyle Font-Size="Smaller" HorizontalAlign="Center" VerticalAlign="Middle" />
                    </asp:BoundField>
                    <asp:BoundField DataField="f_amount" HeaderText="完成數量">
                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                    <ItemStyle Font-Size="Smaller" HorizontalAlign="Center" VerticalAlign="Middle" />
                    </asp:BoundField>
                    <asp:BoundField DataField="notin" HeaderText="入庫數量">
                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                    <ItemStyle Font-Size="Smaller" HorizontalAlign="Center" VerticalAlign="Middle" />
                    </asp:BoundField>
                    <asp:BoundField DataField="rawstatus" HeaderText="料況">
                    <ItemStyle Font-Size="Smaller" HorizontalAlign="Center" VerticalAlign="Middle" />
                    </asp:BoundField>
                    <asp:BoundField DataField="whours" HeaderText="總工時(分)">
                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                    <ItemStyle Font-Size="Smaller" HorizontalAlign="Center" VerticalAlign="Middle" />
                    </asp:BoundField>
                    <asp:BoundField DataField="action" HeaderText="動作">
                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                    <ItemStyle Font-Size="Smaller" HorizontalAlign="Center" VerticalAlign="Middle" />
                    </asp:BoundField>
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
    </form>
</body>
</html>
