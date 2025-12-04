Imports System.Data
Imports System.Data.SqlClient
Partial Public Class WebForm3
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public conn As New SqlConnection
    Public SqlCmd As String
    Public perms As String
    Public ds As New DataSet
    Public dr As SqlDataReader

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        perms = CommUtil.GetAssignRight("ac000", Session("s_id"))
        If (Not IsPostBack) Then
            SqlCmd = "Select areacode,areadesc From dbo.[branch]"
            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
            If (dr.HasRows) Then
                DDLArea.Items.Clear()
                DDLArea.Items.Add("篩選區域")
                Do While (dr.Read())
                    DDLArea.Items.Add(dr(0) & " " & dr(1))
                Loop
            End If
            dr.Close()
            conn.Close()

            SqlCmd = "Select deptcode,deptdesc From dbo.[dept]"
            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
            If (dr.HasRows) Then
                DDLDept.Items.Clear()
                DDLDept.Items.Add("篩選部門")
                Do While (dr.Read())
                    DDLDept.Items.Add(dr(0) & " " & dr(1))
                Loop
            End If
            dr.Close()
            conn.Close()
        End If
        GetUserData()
    End Sub

    Sub GetUserData()
        Dim SelectCmd As String
        Dim rule As String
        rule = ""
        If (DDLArea.SelectedIndex <> 0 And DDLDept.SelectedIndex <> 0) Then
            rule = "where branch='" & Split(DDLArea.SelectedValue, " ")(0) & "' and grp='" & Split(DDLDept.SelectedValue, " ")(0) & "' "
        ElseIf (DDLArea.SelectedIndex <> 0) Then
            rule = "where branch='" & Split(DDLArea.SelectedValue, " ")(0) & "' "
        ElseIf (DDLDept.SelectedIndex <> 0) Then
            rule = "where grp='" & Split(DDLDept.SelectedValue, " ")(0) & "' "
        End If
        SelectCmd = "SELECT T0.id,T0.name,T0.grp,T0.email ,convert(varchar(11),T0.ttl,120) as ttl ,T0.area,T0.denyf,T0.id ,T0.id,T0.position " &
                "FROM dbo.[User] T0 " & rule &
                "order by T0.branch,T0.grp,T0.id"
        'MsgBox(SelectCmd)
        'Dim da1 As New SqlDataAdapter(SelectCmd, conn)
        'da1.Fill(ds)
        ds = CommUtil.SelectLocalSqlUsingDataSet(ds, SelectCmd, conn)
        conn.Close()
        usergv.DataSource = ds
        usergv.DataBind()
    End Sub

    Protected Sub usergv_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles usergv.RowDataBound
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            Dim Hyper, Hyper1 As New HyperLink
            Hyper.Text = e.Row.Cells(8).Text
            If (InStr(perms, "m")) Then
                Hyper.Enabled = True
            Else
                Hyper.Enabled = False
            End If
            Hyper.NavigateUrl = "usermodify.aspx?id=" & e.Row.Cells(0).Text
            e.Row.Cells(8).Controls.Add(Hyper)

            Hyper1.Text = e.Row.Cells(9).Text
            If (InStr(perms, "m")) Then
                Hyper1.Enabled = True
            Else
                Hyper1.Enabled = False
            End If
            'Hyper1.NavigateUrl = "rolemodify.aspx?id=" & e.Row.Cells(0).Text
            Hyper1.NavigateUrl = "rolesetup.aspx?id=" & e.Row.Cells(0).Text
            e.Row.Cells(9).Controls.Add(Hyper1)

            If (e.Row.Cells(6).Text = 1) Then
                e.Row.Cells(6).Text = "可"
            Else
                e.Row.Cells(6).Text = "不可"
            End If
            If (e.Row.Cells(7).Text = 1) Then
                e.Row.Cells(7).Text = "停用"
                e.Row.Cells(7).ForeColor = Drawing.Color.Red
            Else
                e.Row.Cells(7).Text = "正常"
            End If
            If (e.Row.Cells(3).Text = "NA") Then
                e.Row.Cells(3).Text = ""
            End If
        End If
    End Sub

    Protected Sub DDLArea_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DDLArea.SelectedIndexChanged

    End Sub

    Protected Sub DDLDept_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DDLDept.SelectedIndexChanged

    End Sub
End Class