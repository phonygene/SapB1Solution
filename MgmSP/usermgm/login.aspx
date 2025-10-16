<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="login.aspx.vb" Inherits="MgmSP.login" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
&nbsp;<asp:Label ID="Label3" runat="server" Text="Server" Visible="False"></asp:Label>
        &nbsp;
        <asp:TextBox ID="ServerText" runat="server" AutoPostBack="True" Visible="False">.\SQLEXPRESS2008R2</asp:TextBox>
        <br />
    
        <asp:Label ID="Label4" runat="server" Text="資料庫" Visible="False"></asp:Label>
&nbsp;<asp:DropDownList ID="DDLServer" runat="server" Width="170px" 
            AutoPostBack="True" Visible="False">
        </asp:DropDownList>
        <br />
    
        <asp:Label ID="Label5" runat="server" Text="機種"></asp:Label>
        &nbsp;&nbsp;&nbsp;&nbsp;
        <asp:DropDownList ID="DDLWhs" runat="server" Width="170px">
        </asp:DropDownList>
        <br />
    
        <asp:Label ID="Label1" runat="server" Text="帳號"></asp:Label>
&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:TextBox ID="idtxt" runat="server"></asp:TextBox>
        <br />
        <asp:Label ID="Label2" runat="server" Text="密碼"></asp:Label>
&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:TextBox ID="pwdtxt" runat="server" TextMode="Password"></asp:TextBox>
        <br />
        <asp:Label ID="errmsg" runat="server" ForeColor="#FF3300"></asp:Label>
        <br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Button ID="loginbtn" runat="server" Text="登錄" />
    
    </div>
    </form>
</body>
</html>
