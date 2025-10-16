Imports System.Data
Imports System.Data.SqlClient
Public Class ShowData
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public connsap As New SqlConnection
    Public SqlCmd As String
    Public drsap As SqlDataReader
    Public ds As New DataSet
    Public indexpage As Integer
    Public preurl, dtype As String
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim bomcode, tempurl, wo As String
        dtype = Request.QueryString("dtype")
        indexpage = Request.QueryString("indexpage")
        tempurl = Request.QueryString("orgurl")
        preurl = Replace(tempurl, "*", "&")
        CreateHyperMenu()
        If (dtype = "bom") Then
            bomcode = Request.QueryString("bomcode")
            CreateInfo(bomcode)
            CommUtil.ShowBomData(gv1, bomcode)
        ElseIf (dtype = "wo") Then
            wo = Request.QueryString("wo")
            CreateInfo(wo)
            CommUtil.ShowWoData(gv1, wo)
        End If
    End Sub

    Protected Sub gv1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gv1.RowDataBound
        Dim realindex As Integer
        'Dim Hyper As HyperLink

        If (e.Row.RowType = DataControlRowType.DataRow) Then
            realindex = e.Row.RowIndex + gv1.PageIndex * gv1.PageSize
            e.Row.Attributes.Add("onmouseover", "c=this.style.backgroundColor;this.style.backgroundColor='lightgreen'")
            '設定光棒顏色，當滑鼠 onMouseOver 時驅動
            e.Row.Attributes.Add("onmouseout", "this.style.backgroundColor=c")
            '當 onMouseOut 也就是滑鼠移開時，要恢復原本的顏色
            e.Row.Cells(0).Text = realindex + 1
            If ((CInt(e.Row.Cells(5).Text) + CInt(e.Row.Cells(7).Text)) < CInt(e.Row.Cells(6).Text)) Then
                e.Row.Cells(5).BackColor = Drawing.Color.Red
            ElseIf (CInt(e.Row.Cells(5).Text) = 0) Then
                e.Row.Cells(5).BackColor = Drawing.Color.Red
            End If

            If (CInt(e.Row.Cells(8).Text) <> 0) Then
                e.Row.Cells(8).BackColor = Drawing.Color.Yellow
            End If
            If (CInt(e.Row.Cells(9).Text) <> 0) Then
                e.Row.Cells(9).BackColor = Drawing.Color.Yellow
            End If
            If (CInt(e.Row.Cells(10).Text) <> 0) Then
                e.Row.Cells(10).BackColor = Drawing.Color.Yellow
            End If
            If (dtype = "wo") Then
                    If (CInt(e.Row.Cells(11).Text) <> CInt(e.Row.Cells(12).Text)) Then
                        e.Row.Cells(13).BackColor = Drawing.Color.Yellow
                    Else
                        e.Row.Cells(13).BackColor = Drawing.Color.LightGreen
                    End If
                End If
            End If
    End Sub
    Sub CreateInfo(keycode As String)
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim bomcode, bomname As String
        If (dtype = "bom") Then
            bomcode = keycode
        ElseIf (dtype = "wo") Then
            SqlCmd = "Select T0.Itemcode " &
            "from OWOR T0 where T0.docnum='" & keycode & "'"
            drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            drsap.Read()
            bomcode = drsap(0)
            drsap.Close()
            connsap.Close()
        End If
        SqlCmd = "Select T0.Itemname from OITM T0 where T0.itemcode='" & bomcode & "'"
        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        drsap.Read()
        bomname = drsap(0)
        drsap.Close()
        connsap.Close()

        tRow = New TableRow()
        tRow.BorderWidth = 5
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Wrap = False
        If (dtype = "bom") Then
            tCell.Text = bomcode & "&nbsp&nbsp&nbsp" & bomname
        ElseIf (dtype = "wo") Then
            tCell.Text = "工單號:" & keycode & "&nbsp&nbsp&nbsp料號:" & bomcode & "&nbsp&nbsp&nbsp" & bomname
        End If
        tCell.Font.Bold = True
        tRow.Cells.Add(tCell)
        InfoT.Rows.Add(tRow)
    End Sub
    Sub CreateHyperMenu()
        Dim Hyper As HyperLink
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim i, j As Integer
        tRow = New TableRow()
        For i = 0 To 1
            tCell = New TableCell()
            tRow.Cells.Add(tCell)
        Next
        Me.HyperMenuT.Rows.Add(tRow)
        j = 0
        Hyper = New HyperLink()
        Hyper.ID = "index"
        Hyper.Text = "首頁"
        Hyper.NavigateUrl = "../index.aspx?smid=index"
        Hyper.BackColor = Drawing.Color.Aqua
        Hyper.Font.Underline = False
        Hyper.Width = 150
        Hyper.Attributes.Add("onmouseover", "c=this.style.backgroundColor;this.style.backgroundColor='AliceBlue'")
        Hyper.Attributes.Add("onmouseout", "this.style.backgroundColor=c")
        Me.HyperMenuT.Rows(0).Cells(j).HorizontalAlign = HorizontalAlign.Center
        Me.HyperMenuT.Rows(0).Cells(j).Controls.Add(Hyper)
        j = j + 1
        Hyper = New HyperLink()
        Hyper.ID = "backpage"
        Hyper.Text = "回前頁"
        Hyper.NavigateUrl = preurl
        Hyper.BackColor = Drawing.Color.Aqua
        Hyper.Font.Underline = False
        Hyper.Width = 150
        Hyper.Attributes.Add("onmouseover", "c=this.style.backgroundColor;this.style.backgroundColor='AliceBlue'")
        Hyper.Attributes.Add("onmouseout", "this.style.backgroundColor=c")
        Me.HyperMenuT.Rows(0).Cells(j).HorizontalAlign = HorizontalAlign.Center
        Me.HyperMenuT.Rows(0).Cells(j).Controls.Add(Hyper)
    End Sub
End Class