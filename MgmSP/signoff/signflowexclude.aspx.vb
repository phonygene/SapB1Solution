Imports System.Data
Imports System.Data.SqlClient
Public Class signflowexclude
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public connsap, conn As New SqlConnection
    Public SqlCmd As String
    Public dr, dr1 As SqlDataReader
    Public ds As New DataSet
    Public DDLFormType As DropDownList
    Public BtnSave As Button
    Public sfid, formstatusindex As Integer
    Public mode, signid, deptexclude As String
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        If (IsPostBack) Then
            signid = ViewState("signid")
            sfid = ViewState("sfid")
            formstatusindex = ViewState("formstatusindex")
        Else
            signid = Request.QueryString("signid")
            sfid = Request.QueryString("sfid")
            formstatusindex = Request.QueryString("formstatusindex")
            If (Request.QueryString("act") = "saveok") Then
                CommUtil.ShowMsg(Me, "已存檔")
            End If
            ViewState("signid") = signid
            ViewState("sfid") = sfid
            ViewState("formstatusindex") = formstatusindex
        End If
        backsignsetup.NavigateUrl = "~/signoff/signoffsetup.aspx?smid=sg&smode=3&mode=select&formstatusindex=" & formstatusindex & "&sfid=" & sfid
        FTCreate()
    End Sub
    Sub FTCreate()
        Dim tRow As TableRow
        Dim tCell As TableCell
        Dim i, count As Integer
        Dim label1 As Label
        Dim BtnSave As Button
        Dim cChkBox As CheckBox

        SqlCmd = "Select T0.deptcode,T0.deptdesc from dbo.[dept] T0 order by T0.deptcode"
        ds = CommUtil.SelectLocalSqlUsingDataSet(ds, SqlCmd, conn)
        conn.Close()
        count = ds.Tables(0).Rows.Count

        tRow = New TableRow()
        tRow.HorizontalAlign = HorizontalAlign.Center
        tRow.BackColor = Drawing.Color.DeepSkyBlue
        tRow.Font.Bold = True
        tCell = New TableCell()
        tCell.Text = "簽核者"
        tRow.Cells.Add(tCell)

        tCell = New TableCell()
        tCell.Text = "部門排除"
        tRow.Cells.Add(tCell)

        DET.Rows.Add(tRow)

        tRow = New TableRow()
        tRow.HorizontalAlign = HorizontalAlign.Center
        tRow.BackColor = Drawing.Color.LightGreen
        'tRow.Font.Bold = True
        tCell = New TableCell()
        label1 = New Label()
        label1.ID = "label_" & signid
        label1.Text = signid & "<br>"
        tCell.Controls.Add(label1)
        BtnSave = New Button()
        BtnSave.ID = "btn_save"
        BtnSave.Text = "儲存"
        tCell.Controls.Add(BtnSave)
        AddHandler BtnSave.Click, AddressOf BtnSave_Click

        tRow.Cells.Add(tCell)
        tCell = New TableCell()
        For i = 0 To count - 1
            cChkBox = New CheckBox
            cChkBox.ID = "chk_" & ds.Tables(0).Rows(i)("deptcode") & "_" & ds.Tables(0).Rows(i)("deptdesc")
            cChkBox.Text = ds.Tables(0).Rows(i)("deptcode") & "   " & ds.Tables(0).Rows(i)("deptdesc") & "<br>"
            SqlCmd = "select count(*) from [dbo].[@XSDET] where uid='" & signid & "' and deptcode='" & ds.Tables(0).Rows(i)("deptcode") & "' and sfid=" & sfid
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                dr.Read()
                If (dr(0) <> 0) Then
                    cChkBox.Checked = True
                End If
            End If
            dr.Close()
            connsap.Close()
            tCell.Controls.Add(cChkBox)
        Next
        tCell.HorizontalAlign = HorizontalAlign.Left
        tRow.Cells.Add(tCell)
        DET.Rows.Add(tRow)
        ds.Reset()
    End Sub
    Protected Sub BtnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim i, count As Integer
        Dim cChkBox As CheckBox
        Dim chkid As String

        SqlCmd = "delete from dbo.[@XSDET] where sfid=" & sfid & " and uid='" & signid & "'"
        If (CommUtil.SqlSapExecute("del", SqlCmd, connsap)) Then

        Else
            CommUtil.ShowMsg(Me, "新增前,刪除失敗")
        End If
        connsap.Close()

        SqlCmd = "Select T0.deptcode,T0.deptdesc from dbo.[dept] T0 order by T0.deptcode"
        ds = CommUtil.SelectLocalSqlUsingDataSet(ds, SqlCmd, conn)
        conn.Close()
        count = ds.Tables(0).Rows.Count
        For i = 0 To count - 1
            cChkBox = New CheckBox
            chkid = "chk_" & ds.Tables(0).Rows(i)("deptcode") & "_" & ds.Tables(0).Rows(i)("deptdesc")
            If (CType(DET.FindControl(chkid), CheckBox).Checked) Then
                SqlCmd = "insert into dbo.[@XSDET] (sfid,uid,deptcode) " &
                        "values(" & sfid & ",'" & signid & "','" & ds.Tables(0).Rows(i)("deptcode") & "')"
                CommUtil.SqlSapExecute("ins", SqlCmd, connsap)
                connsap.Close()
            End If
        Next
        Response.Redirect("~/signoff/signflowexclude.aspx?smid=sg&smode=3&sfid=" & sfid & "&signid=" & signid & "&formstatusindex=" & formstatusindex & "&act=saveok")
    End Sub
End Class