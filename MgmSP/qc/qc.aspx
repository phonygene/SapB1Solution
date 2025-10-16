<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/MySite1.Master" CodeBehind="qc.aspx.vb" Inherits="MgmSP.qc" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
        <asp:DropDownList ID="DDLFun" runat="server" AutoPostBack="True" Width="183px">
            <asp:ListItem>請選擇操作模式</asp:ListItem>
            <asp:ListItem>建立-料號檢驗項目操作</asp:ListItem>
            <asp:ListItem>查詢-料號檢驗項目操作</asp:ListItem>
            <asp:ListItem>建立-IQC檢驗單操作</asp:ListItem>
            <asp:ListItem>查詢-IQC檢驗單操作</asp:ListItem>
        </asp:DropDownList>
        <asp:Table ID="FTIMC" runat="server" BackColor="#00CCFF" GridLines="Horizontal" Width="100%">
        </asp:Table>
        <asp:Table ID="FTIMS" runat="server" BackColor="#00CCFF" GridLines="Horizontal" Width="100%">
        </asp:Table>
        <asp:Table ID="FTIQCC" runat="server" BackColor="#00CCFF" GridLines="Horizontal" Width="100%">
        </asp:Table>
        <asp:Table ID="FTIQCS" runat="server" BackColor="#00CCFF" GridLines="Horizontal" Width="100%">
        </asp:Table>
        <asp:GridView ID="gv1" runat="server" Width="100%" AutoGenerateColumns="False">
            <AlternatingRowStyle BackColor="#FFFFCC" />
            <HeaderStyle BackColor="#507CD1" BorderStyle="Double" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" />
        </asp:GridView>
</asp:Content>
