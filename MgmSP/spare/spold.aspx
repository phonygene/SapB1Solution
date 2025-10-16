<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/MySite1.Master" CodeBehind="spold.aspx.vb" Inherits="MgmSP.spold" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <asp:DropDownList ID="FTDDL" runat="server" AutoPostBack="True">
        <asp:ListItem>未列帳料件顯示</asp:ListItem>
        <asp:ListItem>未列帳主檔資料建立</asp:ListItem>
    </asp:DropDownList>
    <asp:Table ID="FilterT" runat="server" BackColor="#00CCFF" Width="100%">
    </asp:Table>
    <asp:GridView ID="gv1" runat="server" Width="100%">
        <AlternatingRowStyle BackColor="#FFFFCC" />
        <HeaderStyle BackColor="#507CD1" BorderStyle="Double" />
    </asp:GridView>
    <asp:Table ID="AddT" runat="server">
    </asp:Table>
</asp:Content>
