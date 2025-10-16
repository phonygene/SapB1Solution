Public Class invalid
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim info As String
        Dim tCell As TableCell
        Dim tRow As TableRow
        info = Request.QueryString("info")
        tRow = New TableRow()
        tCell = New TableCell()
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = info
        tCell.Font.Size = 32
        tRow.Cells.Add(tCell)
        InfoT.Rows.Add(tRow)
        Session.Clear()
        Session.Abandon()
    End Sub

End Class