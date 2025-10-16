<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/MySite1.Master" CodeBehind="cncmain.aspx.vb" Inherits="MgmSP.cncmain" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
<!--    <p> -->
        <asp:Table ID="CncAddT" width="100%" runat="server" BackColor="#00CCFF" Font-Size="Smaller" GridLines="Both">
        </asp:Table>
        <asp:Table ID="CncFilterT" width="100%" runat="server" BackColor="#FF9900" BorderColor="#33CCFF" Font-Size="Smaller" GridLines="Both">
        </asp:Table>
        <asp:GridView ID="gv1" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" Font-Size="Smaller" HorizontalAlign="Center" Width="100%" AllowPaging="True" PageSize="14">
            <AlternatingRowStyle BackColor="#FFFFCC" />
            <Columns>
                <asp:BoundField DataField="pseq" HeaderText="順序" >
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:BoundField>
                <asp:BoundField DataField="sappo" HeaderText="採購單" >
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:BoundField>
                <asp:BoundField DataField="num" HeaderText="加工單" >
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:BoundField>
                <asp:BoundField DataField="cdate" HeaderText="建單日期" DataFormatString="{0:yyyy/MM/dd}" >
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:BoundField>
                <asp:BoundField DataField="notfinish" HeaderText="未完/總數" >
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:BoundField>
                <asp:BoundField DataField="vender" HeaderText="廠商" >
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:BoundField>
                <asp:BoundField DataField="cncsta" HeaderText="加工狀態" >
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:BoundField>
                <asp:BoundField DataField="psta" HeaderText="後處裡" >
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:BoundField>
                <asp:BoundField DataField="ship_date" HeaderText="預交日期" DataFormatString="{0:yyyy/MM/dd}" >
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:BoundField>
                <asp:BoundField DataField="comm" HeaderText="備註" />
                <asp:BoundField DataField="rawsta" HeaderText="素材狀態" >
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:BoundField>
                <asp:BoundField DataField="action" HeaderText="動作" >
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:BoundField>
                <asp:BoundField DataField="del" HeaderText="刪除" >
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:BoundField>
            </Columns>
            <EditRowStyle BackColor="#2461BF" />
            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" BorderStyle="Double" HorizontalAlign="Center" VerticalAlign="Middle" />
            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
            <RowStyle BackColor="#EFF3FB" />
            <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
            <SortedAscendingCellStyle BackColor="#F5F7FB" />
            <SortedAscendingHeaderStyle BackColor="#6D95E1" />
            <SortedDescendingCellStyle BackColor="#E9EBEF" />
            <SortedDescendingHeaderStyle BackColor="#4870BE" />
        </asp:GridView>
 <!--   </p> -->
    <p>
    </p>
</asp:Content>
