<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/MySite1.Master" CodeBehind="addsapid.aspx.vb" Inherits="MgmSP.addsapid" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <asp:Panel ID="Panel1" runat="server" BackColor="#FFCC00" Width="350px">
        <asp:Label ID="Label1" runat="server" Text="SAP帳號"></asp:Label>
        &nbsp;
        <asp:TextBox ID="sapidtxt" runat="server"></asp:TextBox>
        <br />
        <asp:Label ID="Label2" runat="server" Text="SAP密碼"></asp:Label>
        &nbsp;
        <asp:TextBox ID="sappwdtxt" runat="server" TextMode="Password"></asp:TextBox>
        <br />
        <asp:Label ID="errmsg" runat="server" ForeColor="#FF3300"></asp:Label>
        <br />
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Button ID="addbtn" runat="server" Text="設定" />
    </asp:Panel>
</asp:Content>
