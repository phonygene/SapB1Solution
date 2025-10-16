<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/MySite1.Master" CodeBehind="freport.aspx.vb" Inherits="MgmSP.freport" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
        <asp:Table ID="FT" runat="server" BackColor="#33CCFF" Width="100%">
        </asp:Table>
        <asp:GridView ID="gv1" runat="server" Width="100%" AutoGenerateColumns="False">
            <AlternatingRowStyle BackColor="#FFFFCC" />
            <FooterStyle BackColor="#FFFF99" />
            <HeaderStyle BorderColor="#507CD1" BorderStyle="Double" />
        </asp:GridView>
        <br />
</asp:Content>
