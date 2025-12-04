<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/MySite1.Master" CodeBehind="signofftool.aspx.vb" Inherits="MgmSP.signofftool" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <asp:Label ID="Label3" runat="server" Text="簽核表欲刪除ID"></asp:Label>
    <asp:DropDownList ID="DDLDelSIgnUser" runat="server" Width="200px" AutoPostBack="True">
    </asp:DropDownList>
    <asp:Label ID="Label1" runat="server" Text="&amp;nbsp&amp;nbsp&amp;nbsp&amp;nbsp&amp;nbsp&amp;nbsp簽核表欲替代ID"></asp:Label>
    <asp:DropDownList ID="DDLReplaceSignUser" runat="server" Width="200px">
    </asp:DropDownList>
    <asp:Label ID="Label2" runat="server" Text="&amp;nbsp&amp;nbsp&amp;nbsp&amp;nbsp&amp;nbsp&amp;nbsp"></asp:Label>
    <asp:Button ID="BtnSignAnalysis" runat="server" Text="簽核需異動分析" />
    <asp:Button ID="BtnSignModify" runat="server" Text="更新簽核異動" OnClientClick="return confirm('要更新嗎')" />
    <br />
    <asp:ListBox ID="MesLB" runat="server" Height="400px" Width="1029px"></asp:ListBox>
    <br />
</asp:Content>
