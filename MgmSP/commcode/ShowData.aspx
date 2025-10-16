<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ShowData.aspx.vb" Inherits="MgmSP.ShowData" %>

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
            <asp:Table ID="InfoT" runat="server" GridLines="Both" Width="100%" BackColor="#3399FF" ForeColor="White">
            </asp:Table>
            <asp:Table ID="CncAddItemT" runat="server" GridLines="Both" Width="100%" BackColor="#99FF66">
            </asp:Table>
            <asp:GridView ID="gv1" runat="server" BackColor="White" ForeColor="#333333" ShowFooter="True" Width="100%">
                <AlternatingRowStyle BackColor="#FFEADF" ForeColor="#284775" />
                <FooterStyle BackColor="#99CCFF" BorderStyle="None" />
                <HeaderStyle BackColor="#99CCFF" BorderStyle="Solid" />
            </asp:GridView>
        </div>
    </form>
</body>
</html>
