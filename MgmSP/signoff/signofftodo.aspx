<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/MySite1.Master" CodeBehind="signofftodo.aspx.vb" Inherits="MgmSP.signofftodo" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <asp:Table ID="contentT" runat="server" Width="100%">
    </asp:Table>
    <asp:DropDownList ID="DDLFormType" runat="server" AutoPostBack="True">
    </asp:DropDownList>
    <asp:DropDownList ID="DDLInCharge" runat="server" AutoPostBack="True">
    </asp:DropDownList>
    <asp:DropDownList ID="DDLTrace" runat="server" AutoPostBack="True">
    </asp:DropDownList>
    <asp:GridView ID="gv1" runat="server" Width="100%">
        <AlternatingRowStyle BackColor="#FFFFCC" />
    </asp:GridView>
    <asp:Table ID="FT" runat="server" BackColor="#99CCFF" Width="100%">
    </asp:Table>
    <iframe id="iframeContent" frameborder="0" width="100%" height="800px" runat="server"></iframe>
</asp:Content>
