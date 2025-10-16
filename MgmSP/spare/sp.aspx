<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/MySite1.Master" CodeBehind="sp.aspx.vb" Inherits="MgmSP.WebForm2" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <asp:GridView ID="gv1" runat="server" AutoGenerateColumns="False" 
        BackColor="#DEBA84" BorderColor="#DEBA84" BorderWidth="1px" 
        CellPadding="3" BorderStyle="None" CellSpacing="2" Width="100%">
        <RowStyle BackColor="#FFF7E7" ForeColor="#8C4510" />
        <Columns>
            <asp:BoundField DataField="DocNum" HeaderText="銷售單">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="U_F1" HeaderText="工單" HtmlEncode="False" />
            <asp:BoundField DataField="ItemCode" HeaderText="料號" HtmlEncode="False" />
            <asp:BoundField DataField="ItemName" HeaderText="說明" HtmlEncode="False" />
            <asp:BoundField DataField="Quantity" DataFormatString="{0:#,##0}" 
                HeaderText="需量" HtmlEncode="False">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="OpenCreQty" DataFormatString="{0:#,##0}" 
                HeaderText="未領" HtmlEncode="False">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="Price" DataFormatString="{0:#,##0}" HeaderText="單價" 
                HtmlEncode="False">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="WhsCode" HeaderText="倉庫" HtmlEncode="False">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="OnHand" DataFormatString="{0:#,##0}" HeaderText="本庫" 
                HtmlEncode="False">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="IsCommited" DataFormatString="{0:#,##0}" 
                HeaderText="本需" HtmlEncode="False">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="OnOrder" DataFormatString="{0:#,##0}" 
                HeaderText="本供" HtmlEncode="False">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="OnHand1" DataFormatString="{0:#,##0}" 
                HeaderText="它庫" HtmlEncode="False">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="IsCommited1" DataFormatString="{0:#,##0}" 
                HeaderText="它需" HtmlEncode="False">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="OnOrder1" DataFormatString="{0:#,##0}" 
                HeaderText="它供" HtmlEncode="False">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="Status" HeaderText="不足">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
        </Columns>
        <FooterStyle BackColor="#F7DFB5" ForeColor="#8C4510" />
        <PagerStyle ForeColor="#8C4510" 
            HorizontalAlign="Center" />
        <SelectedRowStyle BackColor="#738A9C" ForeColor="White" Font-Bold="True" />
        <HeaderStyle BackColor="#A55129" Font-Bold="True" ForeColor="White" />
        <AlternatingRowStyle BackColor="#EAE6E1" />
    </asp:GridView>
</asp:Content>
