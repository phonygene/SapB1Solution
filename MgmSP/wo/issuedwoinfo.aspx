<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/MySite1.Master" CodeBehind="issuedwoinfo.aspx.vb" Inherits="MgmSP.issuedwoinfo" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <asp:Table ID="FT" runat="server" Height="30px">
        </asp:Table>
    <!--<p>-->
        <asp:GridView ID="gv1" runat="server" AutoGenerateColumns="False" 
            CellPadding="4" ForeColor="#333333" AllowPaging="True" AllowSorting="True" 
            BorderStyle="Solid" Width="100%">
            <PagerSettings Mode="NumericFirstLast" />
            <RowStyle BackColor="#EFF3FB" />
            <Columns>
                <asp:BoundField DataField="num" HeaderText="序號">
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Font-Size="Smaller" />
                </asp:BoundField>
                <asp:BoundField DataField="modocnum" HeaderText="MO">
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Font-Size="Smaller" />
                </asp:BoundField>
                <asp:BoundField DataField="wodocnum" HeaderText="WO">
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Font-Size="Smaller" />
                </asp:BoundField>
                <asp:BoundField DataField="issueamount" HeaderText="數量">
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Font-Size="Smaller" />
                </asp:BoundField>
                <asp:BoundField DataField="act" HeaderText="動作" >
                <ItemStyle Font-Size="Smaller" />
                </asp:BoundField>
                <asp:BoundField DataField="wsn" HeaderText="工單號">
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Font-Size="Smaller" />
                </asp:BoundField>
                <asp:BoundField DataField="cus_name" HeaderText="客戶">
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Font-Size="Smaller" />
                </asp:BoundField>
                <asp:BoundField DataField="itemcode" HeaderText="料號">
                <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" Font-Size="Smaller" />
                </asp:BoundField>
                <asp:BoundField DataField="itemname" HeaderText="說明">
                <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" Wrap="True" Width="10%" Font-Size="Smaller" />
                </asp:BoundField>
                <asp:BoundField DataField="req_set" HeaderText="需量">
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Font-Size="Smaller" />
                </asp:BoundField>
                <asp:BoundField DataField="build_date" HeaderText="建單日期" 
                    DataFormatString="{0:d}">
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="True" Font-Size="Smaller" />
                </asp:BoundField>
                <asp:BoundField DataField="req_date" HeaderText="需求日期" DataFormatString="{0:d}">
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="True" Font-Size="Smaller" />
                </asp:BoundField>
                <asp:BoundField DataField="ownername" HeaderText="建立者">
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Font-Size="Smaller" />
                </asp:BoundField>
                <asp:BoundField DataField="issue_count" HeaderText="正領/未領">
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Font-Size="Smaller" />
                </asp:BoundField>
                <asp:BoundField DataField="stat" HeaderText="狀態">
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="5%" 
                    Wrap="False" Font-Size="Smaller" />
                </asp:BoundField>
                <asp:BoundField DataField="upd_date" HeaderText="更新日期" DataFormatString="{0:d}" NullDisplayText="&quot;&quot;">
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="False" Font-Size="Smaller" />
                </asp:BoundField>
                <asp:BoundField DataField="comm" HeaderText="備註" HtmlEncode="False">
                <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" Font-Size="Smaller" />
                </asp:BoundField>
                <asp:BoundField DataField="rtn" HeaderText="退領">
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Font-Size="Smaller" />
                </asp:BoundField>
                <asp:BoundField DataField="del" HeaderText="刪除">
                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Font-Size="Smaller" />
                </asp:BoundField>
            </Columns>
            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" 
                HorizontalAlign="Center" VerticalAlign="Middle" Font-Size="Small" />
            <EditRowStyle BackColor="#2461BF" />
            <AlternatingRowStyle BackColor="White" />
        </asp:GridView>
    <!--</p>-->
    <!--<p>
    </p>-->
</asp:Content>
