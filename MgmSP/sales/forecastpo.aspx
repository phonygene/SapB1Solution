<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="forecastpo.aspx.vb" Inherits="MgmSP.forecastpo" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
            <asp:Table ID="HyperMenuT" runat="server">
            </asp:Table>
            <asp:RadioButtonList ID="MachineOption" runat="server" RepeatDirection="Horizontal" AutoPostBack="True">
                <asp:ListItem Selected="True">AOI機型</asp:ListItem>
                <asp:ListItem>ICT機型</asp:ListItem>
            </asp:RadioButtonList>
            <asp:Table ID="FT" runat="server" BackColor="#33CCFF" Width="100%">
            </asp:Table>
            <asp:Table ID="infoT" runat="server" BackColor="#FFCC66" Width="100%">
            </asp:Table>
            <asp:Table ID="ListT" runat="server" Width="100%">
            </asp:Table>
            <asp:Table ID="UpdT" runat="server">
            </asp:Table>
    </form>
</body>
</html>
