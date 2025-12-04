<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/MySite1.Master" CodeBehind="leave.aspx.vb" Inherits="MgmSP.leave" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <asp:DropDownList ID="FTDDL" runat="server" AutoPostBack="True">
        <asp:ListItem>今日請假人員</asp:ListItem>
        <asp:ListItem>明日請假人員</asp:ListItem>
        <asp:ListItem>以日期查詢請假人員</asp:ListItem>
        <asp:ListItem>填寫請假單</asp:ListItem>
    </asp:DropDownList>
    <asp:Table ID="FilterT" runat="server" BackColor="#00CCFF" Width="100%">
    </asp:Table>
    <asp:Table ID="AddT" runat="server" BackColor="#00CCFF" Width="100%">
    </asp:Table>
    <asp:GridView ID="gv1" runat="server" Width="100%">
        <AlternatingRowStyle BackColor="#FFFFCC" />
        <HeaderStyle BackColor="#507CD1" BorderStyle="Double" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" />
    </asp:GridView>
    <asp:Table ID="UpdT" runat="server">
    </asp:Table>
</asp:Content>
