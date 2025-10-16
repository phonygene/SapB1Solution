<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/MySite1.Master" CodeBehind="molist.aspx.vb" Inherits="MgmSP.molist" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <asp:Table ID="FilterT" runat="server" BackColor="#FF9933" Width="100%">
    </asp:Table>
    <asp:GridView ID="gv1" runat="server" AutoGenerateColumns="False" BorderStyle="Solid" 
        CellPadding="4" Height="10px" AllowSorting="True" 
        Width="100%" HorizontalAlign="Justify"
        AllowPaging="True" ForeColor="#333333" PageSize="15">
        <RowStyle BackColor="#EFF3FB" Width="1px" />
        <Columns>
            <asp:BoundField DataField="docnum" HeaderText="SAP">
            <ItemStyle HorizontalAlign="Center" Height="12" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="wsn" HeaderText="自訂號" HtmlEncode="False" NullDisplayText="NA">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="getpo" HeaderText="訂單狀態">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="cus_name" HeaderText="客戶">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="company" HeaderText="接單部門">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="model" HeaderText="機型">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="resolution" HeaderText="解析度">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="model_set" HeaderText="台數" DataFormatString="{0:G2}">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="f_set" HeaderText="完工數" NullDisplayText="0">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="ship_set" HeaderText="出貨數" NullDisplayText="0">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="ship_date" HeaderText="預計出貨日" DataFormatString="{0:d}" HtmlEncode="False">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="cdate" HeaderText="開單日" DataFormatString="{0:d}" HtmlEncode="False">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="camera_brand" HeaderText="相機">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="True" />
            </asp:BoundField>
            <asp:BoundField DataField="f_stat" HeaderText="狀態">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="note" HeaderText="備註">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="mfmes" HeaderText="紀錄" />
            <asp:BoundField DataField="cno" HeaderText="聯絡單">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
        </Columns>
        <EditRowStyle BackColor="#2461BF" />
        <FooterStyle BackColor="#507CD1" ForeColor="White" Font-Bold="True" />
        <PagerStyle ForeColor="White" HorizontalAlign="Center" BackColor="#2461BF" Font-Size="X-Large" />
        <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
        <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" 
            HorizontalAlign="Center" VerticalAlign="Middle" />
        <AlternatingRowStyle BackColor="#F0F1DA" />
        <SortedAscendingCellStyle BackColor="#F5F7FB" />
        <SortedAscendingHeaderStyle BackColor="#6D95E1" />
        <SortedDescendingCellStyle BackColor="#E9EBEF" />
        <SortedDescendingHeaderStyle BackColor="#4870BE" />
    </asp:GridView>
        <script type="text/javascript">
        function showDisplay1(mes) {
            window.open(mes, "resizeable,scrollbar");
        }
        </script>
    <asp:Label ID="Label1" runat="server" BackColor="#00CCFF" Text="Label"></asp:Label>
</asp:Content>
