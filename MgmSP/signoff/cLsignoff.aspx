<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/MySite1.Master" CodeBehind="cLsignoff.aspx.vb" Inherits="MgmSP.cLsignoff" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <!--<p>-->
        <asp:Table ID="FT_0" runat="server" BackColor="#33CCFF" Width="100%">
        </asp:Table>
    <asp:Table ID="HT" runat="server" BackColor="#99FFCC" Width="100%">
    </asp:Table>
    
        <asp:Table ID="HeadT" runat="server" Width="100%" BackColor="#CCFFFF">
            <asp:TableRow runat="server" BorderStyle="Double" BorderWidth="1px">
                <asp:TableCell runat="server" BorderWidth="1px" Width="12.5%" Wrap="False">表單號</asp:TableCell>
                <asp:TableCell runat="server" BackColor="White" BorderWidth="1px" Width="12.5%" HorizontalAlign="Center"></asp:TableCell>
                <asp:TableCell runat="server" BorderWidth="1px" Width="12.5%">申請人</asp:TableCell>
                <asp:TableCell runat="server" BackColor="White" BorderWidth="1px" Width="12.5%" HorizontalAlign="Center"></asp:TableCell>
                <asp:TableCell runat="server" BorderWidth="1px" Width="12.5%">部門</asp:TableCell>
                <asp:TableCell runat="server" BackColor="White" BorderWidth="1px" Width="12.5%" HorizontalAlign="Center"></asp:TableCell>
                <asp:TableCell runat="server" BorderWidth="1px" Width="12.5%">區域</asp:TableCell>
                <asp:TableCell runat="server" BackColor="White" BorderWidth="1px" Width="12.5%" HorizontalAlign="Center"></asp:TableCell>
            </asp:TableRow>
            <asp:TableRow runat="server" BorderStyle="Double" BorderWidth="1px">
                <asp:TableCell runat="server" BorderWidth="1px" Width="12.5%" Wrap="False">表單類別</asp:TableCell>
                <asp:TableCell runat="server" BackColor="White" BorderWidth="1px" Width="12.5%" HorizontalAlign="Center"></asp:TableCell>
                <asp:TableCell runat="server" BorderWidth="1px" Width="12.5%">表單SAP編號</asp:TableCell>
                <asp:TableCell runat="server" BackColor="White" BorderWidth="1px" Width="12.5%" HorizontalAlign="Center"></asp:TableCell>
                <asp:TableCell runat="server" BorderWidth="1px" Width="12.5%">金額</asp:TableCell>
                <asp:TableCell runat="server" BackColor="White" BorderWidth="1px" Width="12.5%" HorizontalAlign="Center"></asp:TableCell>
                <asp:TableCell runat="server" BackColor="White" BorderWidth="1px" ColumnSpan="2" Width="12.5%"></asp:TableCell>
            </asp:TableRow>
            <asp:TableRow runat="server" BorderStyle="Solid" BorderWidth="1px">
                <asp:TableCell runat="server" BorderWidth="1px" Width="12.5%">主旨</asp:TableCell>
                <asp:TableCell runat="server" BackColor="White" BorderWidth="1px" ColumnSpan="5" Width="12.5%"></asp:TableCell>
                <asp:TableCell runat="server" BorderWidth="1px" Width="12.5%" Wrap="False">附屬單號</asp:TableCell>
                <asp:TableCell runat="server" BackColor="White" BorderWidth="1px" Width="12.5%" HorizontalAlign="Center"></asp:TableCell>
            </asp:TableRow>
        </asp:Table>
        <asp:Table ID="CommT" runat="server" BackColor="White" Width="100%">
    </asp:Table>
        <asp:Table ID="FT_m" runat="server" BackColor="#33CCFF" Width="100%">
        </asp:Table>
    <asp:Table ID="AddT" runat="server" BackColor="#FFCC66" Width="100%">
    </asp:Table>
    <asp:Table ID="FormLogoTitleT" runat="server" Width="100%">
    </asp:Table>
    <asp:Table ID="ContentT" runat="server" Width="100%">
    </asp:Table>
    <iframe id="iframeContent" frameborder="0" width="100%" height="700px" runat="server"></iframe>
        <asp:Table ID="FT_1" runat="server" BackColor="#33CCFF" Width="100%">
    </asp:Table>
        <asp:Table ID="ItemT" runat="server" BackColor="#FFFFCC" Width="100%">
        </asp:Table>
        <asp:Table ID="SignT" runat="server" Width="100%">
        </asp:Table>
    
    <asp:Table ID="CT" runat="server" BackColor="#33CCFF">
    </asp:Table>
    <script type="text/javascript">
        function showDisplay(mes) {
            window.open(mes, "resizeable,scrollbar");
        }
    </script>
</asp:Content>