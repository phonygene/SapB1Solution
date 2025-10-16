<%@ Page Title="" Language="vb" ASPCompat="true" AutoEventWireup="false" MasterPageFile="~/MySite1.Master" CodeBehind="iqc.aspx.vb" Inherits="MgmSP.iqc" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
  <!--  <p> -->
        <asp:Table ID="FT" runat="server" BackColor="#00CCFF" GridLines="Horizontal" Width="100%">
        </asp:Table>
        <asp:Table ID="CT" runat="server" BackColor="#FFFFCC" Width="100%">
        </asp:Table>
    <iframe id="iframeContent" frameborder="0" width="100%" height="900px" runat="server"></iframe>
    <br />
<!--</p>-->
</asp:Content>
