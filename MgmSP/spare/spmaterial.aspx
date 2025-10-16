<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/MySite1.Master" CodeBehind="spmaterial.aspx.vb" Inherits="MgmSP.spmaterial" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <asp:DropDownList ID="DDLWhs" runat="server" Height="19px" Width="168px">
        <asp:ListItem>S04_捷豐備品倉</asp:ListItem>
        <asp:ListItem>S05_捷智通備品倉</asp:ListItem>
        <asp:ListItem>C02_AOI</asp:ListItem>
    </asp:DropDownList>
    <asp:Table ID="FT" runat="server" BackColor="#00CCFF" style="margin-bottom: 0px" Width="100%">
    </asp:Table>
    <asp:GridView ID="gv1" runat="server" AllowSorting="True" AutoGenerateColumns="False" Width="100%" ShowFooter="True">
        <AlternatingRowStyle BackColor="#FFFFCC" />
        <FooterStyle BackColor="#FFFF99" />
    </asp:GridView>
</asp:Content>
