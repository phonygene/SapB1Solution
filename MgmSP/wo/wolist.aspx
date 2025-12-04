<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/MySite1.Master" CodeBehind="wolist.aspx.vb" Inherits="MgmSP.wolist" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <asp:DropDownList ID="DDLWoFun" runat="server" AutoPostBack="True" align="left">
        <asp:ListItem>請選擇要執行的功能</asp:ListItem>
    </asp:DropDownList>
    <asp:DropDownList ID="DDLAlter" runat="server" Enabled="False">
        <asp:ListItem>請選擇替代方法</asp:ListItem>
    </asp:DropDownList>
    <asp:Label ID="Label1" runat="server" Text="需料日期" Font-Size="Small"></asp:Label>
    <asp:TextBox ID="reqdate_text" runat="server" Enabled="False" Width="100px"></asp:TextBox>
    <ajaxToolkit:CalendarExtender ID="reqdate_text_CalendarExtender" Format="yyyy/MM/dd" runat="server" TargetControlID="reqdate_text" />
    <asp:Button ID="ExecuteBtn" runat="server" Text="功能執行" Enabled="False" />
    <asp:CheckBox ID="IssueWoCheck" runat="server" Text="開工單Check" AutoPostBack="True" Enabled="False" />
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <asp:CheckBox ID="IssuedAutoCheck" runat="server" Text="自動轉料" />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    <br />
    <asp:GridView ID="gv1" runat="server" AutoGenerateColumns="False" 
        CellPadding="4" ForeColor="#333333" ShowFooter="True" 
        Width="100%">
        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
        <Columns>
            <asp:BoundField DataField="docnum" HeaderText="SAP號">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="itemcode" HeaderText="料號">
            <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="itemname" HeaderText="規格說明">
            <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="status" HeaderText="狀態">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Wrap="False" />
            </asp:BoundField>
            <asp:BoundField DataField="plannedqty" DataFormatString="{0:#,##0}" 
                HeaderText="計畫&lt;br/&gt;數量" HtmlEncode="False">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="warehouse" HeaderText="倉別">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="onhand" DataFormatString="{0:#,##0}" HeaderText="庫存">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="iscommited" DataFormatString="{0:#,##0}" 
                HeaderText="需求">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="onorder" DataFormatString="{0:#,##0}" 
                HeaderText="供給">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="shortage" DataFormatString="{0:#,##0}" 
                HeaderText="不足">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="sysissued" HeaderText="系統&lt;br&gt;已領" HtmlEncode="False" >
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="issuedqty" DataFormatString="{0:#,##0}" 
                HeaderText="SAP&lt;br&gt;已領" HtmlEncode="False">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="rest" DataFormatString="{0:#,##0}" 
                HeaderText="需領&lt;br&gt;剩餘" HtmlEncode="False">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="amount" DataFormatString="{0:#,##0}" 
                HeaderText="欲領&lt;br&gt;數量" HtmlEncode="False">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
            <asp:BoundField DataField="selchk" HeaderText="選擇">
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:BoundField>
        </Columns>
        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" 
            HorizontalAlign="Center" VerticalAlign="Middle" />
        <PagerStyle ForeColor="White" HorizontalAlign="Center" BorderColor="#FF9966" Font-Size="Larger" />
        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
        <HeaderStyle BackColor="#5D7B9D" BorderStyle="Solid" Font-Bold="True" 
            ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" 
            Wrap="False" />
        <EditRowStyle BackColor="#999999" />
        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
    </asp:GridView>
    <br />
</asp:Content>
