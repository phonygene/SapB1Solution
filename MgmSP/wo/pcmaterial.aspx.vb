Imports System.Data
Imports System.Data.SqlClient
Imports AjaxControlToolkit
Imports System.IO
Public Class pcmaterial
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public CommSignOff As New CommSignOff
    Public connsap, conn As New SqlConnection
    Public SqlCmd As String
    Public dr, drsap As SqlDataReader
    Public ds As New DataSet
    Public ScriptManager1 As New ScriptManager
    Public DDLFun As DropDownList
    Public TxtWo As TextBox
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        'permssg100 = CommUtil.GetAssignRight("sg100", Session("s_id"))

        Page.Form.Controls.Add(ScriptManager1)
        FTCreate()
        'FormListDisplay()
    End Sub
    Sub FTCreate()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim Labelx As Label
        tRow = New TableRow()
        tCell = New TableCell()
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Left

        Labelx = New Label
        Labelx.ID = "label_1"
        Labelx.Text = "指定工單:&nbsp"
        tCell.Controls.Add(Labelx)

        TxtWo = New TextBox
        TxtWo.Width = 60
        tCell.Controls.Add(TxtWo)

        Labelx = New Label
        Labelx.ID = "label_2"
        Labelx.Text = "&nbsp&nbsp"
        tCell.Controls.Add(Labelx)

        DDLFun = New DropDownList
        DDLFun.ID = "ddl_fun"
        DDLFun.Width = 250
        DDLFun.AutoPostBack = True

        DDLFun.Items.Clear()
        DDLFun.Items.Add("請選擇料況查詢")
        DDLFun.Items.Add("已發料,但因缺料未發完(現料已足 可再發料)")
        AddHandler DDLFun.SelectedIndexChanged, AddressOf DDLFun_SelectedIndexChanged
        tCell.Controls.Add(DDLFun)

        tRow.Controls.Add(tCell)
        FT.Rows.Add(tRow)
    End Sub
    Protected Sub DDLFun_SelectedIndexChanged(sender As Object, e As EventArgs)
        If (DDLFun.SelectedIndex <> 0) Then
            FormListDisplay(DDLFun.SelectedIndex)
        Else

        End If
        'Response.Redirect("~/signoff/signoff.aspx?smid=sg&smode=1&formtypeindex=" & DDLFormType.SelectedIndex &
        '                  "&formstatusindex=" & DDLFormStatus.SelectedIndex & "&signflowmode=" & CType(FT.FindControl("signflowmode"), RadioButtonList).SelectedIndex)
    End Sub
    Sub FormListDisplay(funnum As Integer)

        ds.Reset()
        SetGridViewStyle()
        If (funnum = 1) Then
            SetForm1ListGridViewFields()
        End If
        SqlCmd = "SELECT T0.[DocNum],T0.U_F16,T2.[ItemCode], T2.[ItemName], T0.[Warehouse], " &
                "T0.[DueDate], T1.[BaseQty]*T0.[CmpltQty]-T1.[IssuedQty] as lackamount, T3.[OnHand] " &
                "FROM OWOR T0 INNER JOIN WOR1 T1 ON T0.DocEntry = T1.DocEntry INNER JOIN OITM T2 ON T1.[ItemCode] = T2.[ItemCode] " &
                "INNER JOIN OITW T3 ON T1.[ItemCode] = T3.[ItemCode] WHERE T1.[IssuedQty] < T1.[BaseQty] * T0.[CmpltQty] And " &
                "T0.[CmpltQty]>0 And T3.[OnHand]>= T1.[BaseQty]*T0.[CmpltQty]-T1.[IssuedQty] And T0.[Status]<>'L'  and " &
                "T0.[Warehouse]='" & Session("usingwhs") & "' And T3.WhsCode='" & Session("usingwhs") & "' ORDER BY T0.[DocNum]"
        ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
        connsap.Close()


        'ds.Tables(0).Columns.Add("action")

        'If (ds.Tables(0).Rows.Count = 0) Then
        '    CommUtil.ShowMsg(Me, "無任何資料")
        'End If
        gv1.DataSource = ds.Tables(0)
        gv1.DataBind()

    End Sub
    Sub SetGridViewStyle()
        gv1.AutoGenerateColumns = False
        'gv1.ShowHeader = True
        'gv1.AllowPaging = True
        'gv1.PageSize = 20
        gv1.PagerStyle.HorizontalAlign = HorizontalAlign.Center
        'gv1.AllowSorting = True
        'gv1.Font.Size = FontSize.Smaller
        'gv1.ForeColor =
        gv1.GridLines = GridLines.Both
        gv1.HeaderStyle.HorizontalAlign = HorizontalAlign.Center
        gv1.FooterStyle.HorizontalAlign = HorizontalAlign.Center
        'gv1.HeaderStyle.BackColor =
        'gv1.RowStyle.BackColor
        'gv1.AlternatingRowStyle.BackColor
        'gv1.HeaderStyle.ForeColor
    End Sub

    Sub SetForm1ListGridViewFields()
        Dim oBoundField As BoundField
        gv1.Columns.Clear()
        oBoundField = New BoundField
        oBoundField.HeaderText = "工單號"
        oBoundField.DataField = "docnum"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "母工單號"
        oBoundField.DataField = "U_F16"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "料號"
        oBoundField.DataField = "itemcode"
        oBoundField.ShowHeader = True
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "說明"
        oBoundField.DataField = "itemname"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:N0}"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "倉庫"
        oBoundField.DataField = "warehouse"
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ShowHeader = True
        oBoundField.HtmlEncode = False
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "未領數量"
        oBoundField.DataField = "lackamount"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "目前庫存"
        oBoundField.DataField = "onhand"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)
    End Sub
End Class