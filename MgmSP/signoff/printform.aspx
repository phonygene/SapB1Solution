<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="printform.aspx.vb" Inherits="MgmSP.printform" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:Table ID="FormLogoTitleT" runat="server" Width="100%">
            </asp:Table>
            <asp:Table ID="contentT" runat="server" Width="100%">
            </asp:Table>
            <asp:Table ID="SignT" runat="server" Width="100%">
            </asp:Table>
        </div>
    </form>
        <script type="text/javascript">
        function showDisplay(mes) {
            window.open(mes, "resizeable,scrollbar");
        }
        </script>
</body>
</html>
