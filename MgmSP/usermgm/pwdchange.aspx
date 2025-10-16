<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/MySite1.Master" CodeBehind="pwdchange.aspx.vb" Inherits="MgmSP.pwdchange" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <asp:Panel ID="Panel1" runat="server" BackColor="#33CCFF" 
        Direction="LeftToRight" Height="167px" HorizontalAlign="Center" Width="408px">
        <asp:Label ID="Label1" runat="server" Text="舊密碼"></asp:Label>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:TextBox ID="oldpwdtxt" runat="server" TextMode="Password"></asp:TextBox>
        <br />
        <asp:Label ID="Label2" runat="server" Text="新密碼"></asp:Label>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:TextBox ID="newpwdtxt" runat="server" TextMode="Password"></asp:TextBox>
        <br />
        <asp:Label ID="Label3" runat="server" Text="新密碼確認"></asp:Label>
        &nbsp;
        <asp:TextBox ID="cnewpwdtxt" runat="server"></asp:TextBox>
        <br />
        <asp:Label ID="errmsg" runat="server" ForeColor="#FF3300"></asp:Label>
        <br />
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Button ID="modifybtn" runat="server" Text="修改" />
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <br />
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    </asp:Panel>
</asp:Content>
