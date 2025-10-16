Imports System.Data
Imports System.Data.SqlClient
Imports AjaxControlToolkit
Imports System.IO
Public Class signoff
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public CommSignOff As New CommSignOff
    Public connsap, conn, connsap1 As New SqlConnection
    Public SqlCmd As String
    Public dr, drsap, dr1 As SqlDataReader
    Public ds As New DataSet
    Public ScriptManager1 As New ScriptManager
    Public DDLFormType, DDLFormStatus As DropDownList
    Public BtnFormAdd, BtnFilter As Button
    Public BtnEmail As Button
    Public sid As String
    Public TxtDocnum, TxtKW As TextBox
    Public act As String

    Protected Sub DDLFormType_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim str() As String
        str = Split(DDLFormType.SelectedValue, " ")
        If (CInt(str(1)) > 70 And CInt(str(1)) < 80) Then
            SqlCmd = "select grp from dbo.[user] where id='" & sid & "'"
            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
            If (dr.HasRows) Then
                dr.Read()
                If (dr(0) <> "PD") Then
                    CommUtil.ShowMsg(Me, "你的部門別不是採購,無法執行此採購單")
                    DDLFormType.SelectedIndex = Request.QueryString("formtypeindex")
                    Exit Sub
                End If
            Else
                CommUtil.ShowMsg(Me, "沒找到id為" & sid & "之資料,請檢查")
                'Exit Function
            End If
            dr.Close()
            conn.Close()
        End If

        Response.Redirect("~/signoff/signoff.aspx?smid=sg&smode=1&formtypeindex=" & DDLFormType.SelectedIndex &
                  "&formstatusindex=" & DDLFormStatus.SelectedIndex & "&signflowmode=" & CType(FT.FindControl("signflowmode"), RadioButtonList).SelectedIndex)
        'MsgBox(DDLFormType.SelectedIndex & "-" & DDLFormStatus.SelectedIndex)
    End Sub

    Protected Sub DDLFormStatus_SelectedIndexChanged(sender As Object, e As EventArgs)
        If (DDLFormStatus.SelectedIndex <> 0) Then
            CType(FT.FindControl("signflowmode"), RadioButtonList).SelectedIndex = 1
        Else
            CType(FT.FindControl("signflowmode"), RadioButtonList).SelectedIndex = 0
        End If
        Response.Redirect("~/signoff/signoff.aspx?smid=sg&smode=1&formtypeindex=" & DDLFormType.SelectedIndex &
                          "&formstatusindex=" & DDLFormStatus.SelectedIndex & "&signflowmode=" & CType(FT.FindControl("signflowmode"), RadioButtonList).SelectedIndex)
    End Sub
    Protected Sub BtnFilter_Click(sender As Object, e As EventArgs)
        If (TxtDocnum.Text <> "") Then
            If (Not IsNumeric(TxtDocnum.Text)) Then
                CommUtil.ShowMsg(Me, "表單號欄位不是數字,請更正")
                Exit Sub
            End If
        End If
        If (TxtDocnum.Text <> "" Or TxtKW.Text <> "") Then
            Response.Redirect("~/signoff/signoff.aspx?smid=sg&smode=1&act=mysearch&txtdocnum=" & TxtDocnum.Text &
                          "&txtkw=" & TxtKW.Text & "&signflowmode=" & CType(FT.FindControl("signflowmode"), RadioButtonList).SelectedIndex)
        Else
            CommUtil.ShowMsg(Me, "表單號或主旨關鍵字須輸入")
        End If

    End Sub

    Protected Sub BtnFormAdd_Click(sender As Object, e As EventArgs)
        Dim sfid As Integer
        Dim str() As String
        If (DDLFormType.SelectedIndex <> 0) Then
            str = Split(DDLFormType.SelectedValue, " ")
            sfid = CInt(str(1))
            Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&actmode=single&act=add&status=A&sfid=" & sfid & "&formstatusindex=1&formtypeindex=" & DDLFormType.SelectedIndex)
        End If
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '一般聯絡單 sfid 1
        '備品聯絡單 sfid 2
        '機台聯絡單 sfid 3
        '採購單 sfid 71
        sid = Session("s_id")
        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        'permssg100 = CommUtil.GetAssignRight("sg100", Session("s_id"))

        Page.Form.Controls.Add(ScriptManager1)
        FTCreate()
        If (Not IsPostBack) Then
            gv1.PageIndex = Request.QueryString("indexpage")
            act = Request.QueryString("act")
            'If (act <> "") Then
            DDLFormType.SelectedIndex = Request.QueryString("formtypeindex")
            DDLFormStatus.SelectedIndex = Request.QueryString("formstatusindex")
            CType(FT.FindControl("signflowmode"), RadioButtonList).SelectedIndex = Request.QueryString("signflowmode")
            If (DDLFormStatus.SelectedIndex <> 0) Then
                CType(FT.FindControl("signflowmode"), RadioButtonList).Enabled = False
            Else
                CType(FT.FindControl("signflowmode"), RadioButtonList).Enabled = True
            End If
            'End If
            If (act = "mysearch") Then
                TxtDocnum.Text = Request.QueryString("txtdocnum")
                TxtKW.Text = Request.QueryString("txtkw")
                SearchDoc()
            Else
                FormListDisplay(1)
            End If
            If (act = "del") Then
                CommUtil.ShowMsg(Me, "已刪除")
            ElseIf (act = "signfinish") Then
                CommUtil.ShowMsg(Me, "簽核已全部完成")
            End If
            If (DDLFormType.SelectedIndex = 0) Then
                BtnFormAdd.Enabled = False
            Else
                BtnFormAdd.Enabled = True
            End If
        End If
        'Dim di As DirectoryInfo
        'di = New DirectoryInfo(HttpContext.Current.Server.MapPath("~/") & "AttachFile\SignOffsFormFiles\ron\49\")
        'Dim fi As FileInfo()
        'Dim fname As String
        'Dim targetPath, localsignoffformdir As String
        'localsignoffformdir = Application("localdir") & "SignOffsFormFiles\ron\49\"
        'targetPath = HttpContext.Current.Server.MapPath("~/") & "AttachFile\SignOffsFormFiles\ron\49\"
        'fi = di.GetFiles("*附檔*")
        'For k = 0 To fi.Length - 1
        '    If (InStr(fi(k).Name, "Stamped") = 0 And InStr(LCase(fi(k).Name), "pdf") <> 0) Then
        '        'MsgBox(fi(k).Name)
        '        CommSignOff.GenPdfStamper1(fi(k).Name, Split(fi(k).Name, ".")(0) & "_Stamped.pdf", targetPath, localsignoffformdir, 1, False) '產生浮水印pdf 並刪除approved pdf
        '    End If
        'Next
        '以下為Test
        'Dim fileNameSign, targetPath As String
        'Dim pdfFiles(1) As String
        'Dim fileNameApproved As String = "6_Approved.pdf"
        'targetPath = HttpContext.Current.Server.MapPath("~/") & "SignOffsFormFiles\71\tedy\"
        'fileNameSign = "6_sign.pdf"
        'pdfFiles(0) = "6.pdf"
        'pdfFiles(1) = fileNameSign '會簽流程的pdf
        'CommUtil.mergePDF(pdfFiles, fileNameApproved, targetPath)
        'CommSignOff.createRSCPDF(48, "Jet門禁磁卡補刷卡單", "kk.pdf", "c:\test\", 16)
        'CommSignOff.createMaterialInOutPDF(2, "料件(物品)收發單", "kkk.pdf", "c:\test\", 50000, "test", "NTD", 51)
        'Dim url = "http://localhost:50601/"
        'CommSignOff.HtmlToPdfGen(url, 28, 1,Application("localdir"))
        'CommSignOff.ToDoListPush(Application("http"))
        'If (errstr <> "") Then
        '    CommUtil.ShowMsg(Me, errstr)
        'End If
        'filename test
        'Dim di As DirectoryInfo
        'di = New DirectoryInfo("C:\sapupload\")
        'Dim fi As FileInfo() = di.GetFiles("1_*")
        'For i = 0 To fi.Length - 1
        '    MsgBox(fi(i).Name)
        'Next
        'Response.Redirect("~/signoff/printform.aspx?docnum=25&sfid=101&usingwhs=C02")
        'CommSignOff.SignOffPush(Application("http"), 1)
    End Sub
    Protected Sub gv1_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gv1.RowDataBound
        Dim realindex As Integer
        Dim Hyper As HyperLink
        Dim sfid As Integer
        Dim drL As SqlDataReader
        Dim connL As New SqlConnection
        Dim rtnfinish As Boolean
        rtnfinish = False
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            If (e.Row.Cells(3).Text = "0") Then
                e.Row.Cells(3).Text = "NA"
                'MsgBox(e.Row.Cells(2).Text)
            End If
            If (e.Row.Cells(4).Text = "") Then
                e.Row.Cells(4).Text = "NA"
            End If
            sfid = e.Row.Cells(1).Text
            realindex = e.Row.RowIndex + gv1.PageIndex * gv1.PageSize
            SqlCmd = "Select T0.sfname from dbo.[@XSFTT] T0 where T0.sfid=" & e.Row.Cells(1).Text
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            e.Row.Cells(1).Text = dr(0)
            dr.Close()
            connsap.Close()
            If (e.Row.Cells(6).Text = "E" Or e.Row.Cells(6).Text = "D" Or e.Row.Cells(6).Text = "B" Or e.Row.Cells(6).Text = "R" Or e.Row.Cells(6).Text = "F" Or e.Row.Cells(6).Text = "T") Then
                e.Row.Cells(7).Text = "NA"
            ElseIf (e.Row.Cells(6).Text = "O") Then
                SqlCmd = "SELECT seq,status FROM dbo.[@XSPWT] where docentry=" & e.Row.Cells(0).Text & " and uid='" & Session("s_id") & "'"
                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                If (dr.HasRows) Then
                    dr.Read()
                    If (dr(1) = 0 Or dr(1) = 1) Then
                        e.Row.Cells(7).Text = "NA"
                    ElseIf (dr(1) = 2 Or dr(1) = 100) Then
                        SqlCmd = "Select T0.status from dbo.[@XSPWT] T0 where signprop=0 and T0.docentry=" & e.Row.Cells(0).Text & " and T0.seq=" & dr(0) + 1
                        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                        If (drL.HasRows) Then
                            drL.Read()
                            If (drL(0) = 1) Then
                                e.Row.Cells(7).Text = "可"
                            Else
                                e.Row.Cells(7).Text = "否"
                            End If
                        Else
                            e.Row.Cells(7).Text = "否"
                        End If
                        drL.Close()
                        connL.Close()
                    Else
                        e.Row.Cells(7).Text = "NA"
                    End If
                Else
                    e.Row.Cells(7).Text = "NA"
                End If
                dr.Close()
                connsap.Close()
            Else
                e.Row.Cells(7).Text = "NA"
            End If

            Hyper = New HyperLink
            Hyper.ID = "hyper_action"
            If (act <> "mysearch") Then
                If (DDLFormStatus.SelectedIndex = 0) Then
                    If (e.Row.Cells(6).Text = "E") Then
                        Hyper.Text = "編輯中"
                    ElseIf (e.Row.Cells(6).Text = "D") Then
                        Hyper.Text = "待送審"
                    ElseIf (e.Row.Cells(6).Text = "B" Or e.Row.Cells(6).Text = "R") Then
                        Hyper.Text = "再送審"
                    ElseIf (e.Row.Cells(6).Text = "F") Then
                        SqlCmd = "SELECT signprop FROM dbo.[@XSPWT] where status=1 and docentry=" & e.Row.Cells(0).Text & " and uid='" & Session("s_id") & "'"
                        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                        If (drL.HasRows) Then
                            drL.Read()
                            If (drL(0) = 1) Then
                                Hyper.Text = "待歸檔"
                            Else
                                Hyper.Text = "待知悉"
                            End If
                        End If
                        drL.Close()
                        connL.Close()
                    ElseIf (e.Row.Cells(6).Text = "T") Then
                        'SqlCmd = "SELECT signprop FROM dbo.[@XSPWT] where status=1 and docentry=" & e.Row.Cells(0).Text & " and uid='" & Session("s_id") & "'"
                        'drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                        'If (drL.HasRows) Then
                        '    drL.Read()
                        '    If (drL(0) = 2) Then
                        Hyper.Text = "待知悉"
                        'End If
                        'End If
                        'drL.Close()
                        'connL.Close()
                    ElseIf (e.Row.Cells(6).Text = "O") Then
                        Hyper.Text = "待覆核"
                    Else
                        Hyper.Text = "顯示"
                    End If
                Else
                    Hyper.Text = "顯示"
                End If
            Else
                If (e.Row.Cells(6).Text = "E") Then
                    Hyper.Text = "編輯中"
                ElseIf (e.Row.Cells(6).Text = "D") Then
                    Hyper.Text = "待送審"
                ElseIf (e.Row.Cells(6).Text = "C") Then
                    Hyper.Text = "已作廢"
                ElseIf (e.Row.Cells(6).Text = "B" Or e.Row.Cells(6).Text = "R") Then
                    Hyper.Text = "再送審"
                Else
                    If (e.Row.Cells(6).Text = "O") Then
                        SqlCmd = "SELECT status,signprop FROM dbo.[@XSPWT] where docentry=" & e.Row.Cells(0).Text & " and uid='" & Session("s_id") & "' order by seq" '為了若有2個簽核人(送審人及歸檔是同一人)
                    Else
                        SqlCmd = "SELECT status,signprop FROM dbo.[@XSPWT] where docentry=" & e.Row.Cells(0).Text & " and uid='" & Session("s_id") & "' order by seq desc"
                    End If
                    drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                    If (drL.HasRows) Then
                        drL.Read()
                        If (drL(0) = 1 And drL(1) = 0) Then
                            Hyper.Text = "待覆核"
                        ElseIf (drL(0) = 1 And drL(1) = 1) Then
                            Hyper.Text = "待歸檔"
                        ElseIf (drL(0) = 1 And drL(1) = 2) Then
                            Hyper.Text = "待知悉"
                        ElseIf (drL(0) = 0) Then
                            Hyper.Text = "未到之簽核"
                        ElseIf (drL(0) = 100 And drL(1) = 0) Then
                            Hyper.Text = "已送審"
                        ElseIf (drL(0) = 2 And drL(1) = 0) Then
                            Hyper.Text = "已覆核"
                        ElseIf (drL(0) = 2 And drL(1) = 1) Then
                            Hyper.Text = "已歸檔"
                        ElseIf (drL(0) = 104 And drL(1) = 2) Then
                            Hyper.Text = "已知悉"
                        Else
                            Hyper.Text = "顯示"
                        End If
                    End If
                    drL.Close()
                    connL.Close()
                End If
            End If
            Dim actmode As String
            actmode = ""
            If (DDLFormStatus.SelectedIndex = 0) Then
                If (CType(FT.FindControl("signflowmode"), RadioButtonList).SelectedIndex = 0) Then
                    If (act <> "mysearch") Then
                        If (e.Row.Cells(6).Text <> "D" And e.Row.Cells(6).Text <> "E" And e.Row.Cells(6).Text <> "B" And e.Row.Cells(6).Text <> "R") Then
                            actmode = "signoff_login"
                        Else
                            actmode = "single"
                        End If
                    Else
                        actmode = "single"
                    End If
                Else
                    actmode = "single"
                End If
            Else
                actmode = "single"
            End If
            Hyper.NavigateUrl = "cLsignoff.aspx?smid=sg&smode=2&status=" & e.Row.Cells(6).Text &
                                "&indexpage=" & gv1.PageIndex & "&docnum=" & e.Row.Cells(0).Text &
                                "&actmode=" & actmode & "&formstatusindex=" & DDLFormStatus.SelectedIndex &
                                "&formtypeindex=" & DDLFormType.SelectedIndex & "&sfid=" & sfid & "&signflowmode=" & Request.QueryString("signflowmode") &
                                "&fromasp=signoff"
            Hyper.Font.Underline = False
            e.Row.Cells(9).Controls.Add(Hyper)
            Dim sfid101process, sfid100process As Boolean
            Dim s101docnum As Long
            Dim s100docnum As Long
            sfid101process = False
            sfid100process = False
            If ((e.Row.Cells(6).Text = "T" Or e.Row.Cells(6).Text = "F") And sfid <> 100 And sfid <> 101) Then
                If (sfid = 23 Or sfid = 24) Then
                    SqlCmd = "Select docnum FROM [dbo].[@XASCH] T0 WHERE T0.[attadoc] ='" & e.Row.Cells(0).Text & "' and sfid=101 and status<>'F' and status<>'T' and status<>'C'"
                    dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, conn)
                    If (dr.HasRows) Then
                        dr.Read()
                        s101docnum = dr(0)
                        sfid101process = True
                    End If
                    dr.Close()
                    conn.Close()
                    If (sfid101process = False) Then
                        'Check 此單是否還有料件要返還
                        SqlCmd = "Select sum(quantity),sum(rtnqty) FROM [dbo].[@XSMLS] T0 WHERE T0.[docentry] =" & e.Row.Cells(0).Text & " And head=0"
                        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, conn)
                        dr.Read()
                        If (dr(0) = dr(1)) Then
                            rtnfinish = True
                        End If
                        dr.Close()
                        conn.Close()
                        If (Not rtnfinish) Then
                            Hyper = New HyperLink
                            Hyper.ID = "hyper_addsignoff"
                            Hyper.NavigateUrl = "~/signoff/cLsignoff.aspx?smid=sg&smode=2&actmode=single&act=add101&status=A&sfid=101" &
                        "&formtypeindex=" & DDLFormType.SelectedIndex & "&formstatusindex=" & DDLFormStatus.SelectedIndex &
                        "&maindocnum=" & ds.Tables(0).Rows(realindex)("docnum") & "&indexpage=" & gv1.PageIndex &
                        "&signflowmode=" & Request.QueryString("signflowmode") & "&fromasp=signoff"
                            Hyper.Text = "新增返還單"
                            Hyper.Font.Underline = False
                            e.Row.Cells(10).Controls.Add(Hyper)
                        Else
                            e.Row.Cells(10).Text = "NA"
                        End If
                    Else
                        e.Row.Cells(10).Text = "返還簽核中(" & s101docnum & ")"
                    End If
                Else
                    SqlCmd = "Select docnum FROM [dbo].[@XASCH] T0 WHERE T0.[attadoc] ='" & e.Row.Cells(0).Text & "' and sfid=100 and status<>'F' and status<>'T' and status<>'C'"
                    dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, conn)
                    If (dr.HasRows) Then
                        dr.Read()
                        s100docnum = dr(0)
                        sfid100process = True
                    End If
                    dr.Close()
                    conn.Close()
                    If (sfid100process = False) Then
                        Hyper = New HyperLink
                        Hyper.ID = "hyper_addsignoff"
                        Hyper.NavigateUrl = "~/signoff/cLsignoff.aspx?smid=sg&smode=2&actmode=single&act=add&status=A&sfid=100" &
                        "&formtypeindex=" & DDLFormType.SelectedIndex & "&formstatusindex=" & DDLFormStatus.SelectedIndex &
                        "&maindocnum=" & ds.Tables(0).Rows(realindex)("docnum") & "&indexpage=" & gv1.PageIndex &
                        "&signflowmode=" & Request.QueryString("signflowmode") & "&fromasp=signoff"
                        Hyper.Text = "新增補充加簽"
                        Hyper.Font.Underline = False
                        e.Row.Cells(10).Controls.Add(Hyper)
                    Else
                        e.Row.Cells(10).Text = "補充簽核中(" & s100docnum & ")"
                    End If
                End If
            Else
                e.Row.Cells(10).Text = "NA"
            End If


            If (e.Row.Cells(6).Text = "E") Then
                e.Row.Cells(6).Text = "編輯中"
            ElseIf (e.Row.Cells(6).Text = "D") Then
                e.Row.Cells(6).Text = "底稿"
            ElseIf (e.Row.Cells(6).Text = "O") Then
                e.Row.Cells(6).Text = "簽核中"
            ElseIf (e.Row.Cells(6).Text = "F") Then
                e.Row.Cells(6).Text = "簽核完成"
            ElseIf (e.Row.Cells(6).Text = "C") Then
                e.Row.Cells(6).Text = "作廢"
            ElseIf (e.Row.Cells(6).Text = "R") Then
                e.Row.Cells(6).Text = "抽回"
            ElseIf (e.Row.Cells(6).Text = "B") Then
                e.Row.Cells(6).Text = "駁回"
            ElseIf (e.Row.Cells(6).Text = "T") Then
                e.Row.Cells(6).Text = "已歸檔"
            End If
        End If
    End Sub
    Protected Sub gv1_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles gv1.PageIndexChanging
        gv1.PageIndex = e.NewPageIndex
        '以下或可用 Response.Redirect 帶參數方式執行
        If (act = "mysearch") Then
            TxtDocnum.Text = Request.QueryString("txtdocnum")
            TxtKW.Text = Request.QueryString("txtkw")
            SearchDoc()
        Else
            FormListDisplay(1)
        End If
    End Sub
    Sub FTCreate()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim Labelx As Label
        Dim rRBL As RadioButtonList
        tRow = New TableRow()
        tCell = New TableCell()
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Left
        DDLFormType = New DropDownList
        DDLFormType.ID = "ddl_formtype"
        DDLFormType.Width = 250
        DDLFormType.AutoPostBack = True
        SqlCmd = "Select T0.sfname,T0.sfid,T0.sftypenote,deptcode from dbo.[@XSFTT] T0 order by T0.sfid" 'where sfid<=54 order by T0.sfid"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            DDLFormType.Items.Clear()
            DDLFormType.Items.Add("所有簽單 0")
            Do While (dr.Read())
                'If (dr(3) = "") Then
                DDLFormType.Items.Add(dr(0) & " " & dr(1) & " " & dr(2))
                'Else
                'If (InStr(dr(3), Session("grp")) Or InStr(dr(3), Session("branch"))) Then
                'DDLFormType.Items.Add(dr(0) & " " & dr(1) & " " & dr(2))
                '
                'End If
            Loop
        End If
        dr.Close()
        connsap.Close()
        AddHandler DDLFormType.SelectedIndexChanged, AddressOf DDLFormType_SelectedIndexChanged
        tCell.Controls.Add(DDLFormType)

        Labelx = New Label
        Labelx.ID = "label_1"
        Labelx.Text = "&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)

        DDLFormStatus = New DropDownList
        DDLFormStatus.ID = "ddl_formstatus"
        DDLFormStatus.Width = 150
        DDLFormStatus.Items.Clear()
        DDLFormStatus.Items.Add("我的待簽單")
        DDLFormStatus.Items.Add("我的已送審單")
        DDLFormStatus.Items.Add("我的已核准")
        DDLFormStatus.Items.Add("我的已駁回")
        DDLFormStatus.Items.Add("我的待送審單")
        DDLFormStatus.Items.Add("我的作廢單")
        DDLFormStatus.Items.Add("我的待歸檔")
        DDLFormStatus.Items.Add("我的已歸檔")
        DDLFormStatus.Items.Add("我的待知悉")
        DDLFormStatus.Items.Add("我的已知悉")
        DDLFormStatus.AutoPostBack = True
        AddHandler DDLFormStatus.SelectedIndexChanged, AddressOf DDLFormStatus_SelectedIndexChanged
        tCell.Controls.Add(DDLFormStatus)

        Labelx = New Label
        Labelx.ID = "label_2"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        BtnFormAdd = New Button
        BtnFormAdd.ID = "btn_add"
        BtnFormAdd.Text = "簽核單新增"
        BtnFormAdd.Enabled = False
        'BtnFormAdd.OnClientClick = "return confirm('確定要新增嗎')"
        AddHandler BtnFormAdd.Click, AddressOf BtnFormAdd_Click
        tCell.Controls.Add(BtnFormAdd)

        Labelx = New Label
        Labelx.ID = "label_3"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp表單號:"
        tCell.Controls.Add(Labelx)
        TxtDocnum = New TextBox
        TxtDocnum.Width = 60
        tCell.Controls.Add(TxtDocnum)

        Labelx = New Label
        Labelx.ID = "label_4"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp主旨關鍵字:"
        tCell.Controls.Add(Labelx)
        TxtKW = New TextBox
        TxtKW.Width = 120
        tCell.Controls.Add(TxtKW)

        Labelx = New Label
        Labelx.ID = "label_5"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        BtnFilter = New Button
        BtnFilter.ID = "btn_filter"
        BtnFilter.Text = "查詢"
        AddHandler BtnFilter.Click, AddressOf BtnFilter_Click
        tCell.Controls.Add(BtnFilter)
        tRow.Controls.Add(tCell)

        tCell = New TableCell()
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Right
        rRBL = New RadioButtonList
        rRBL.ID = "signflowmode"
        rRBL.Items.Add("連續簽核")
        rRBL.Items.Add("單一簽核")
        rRBL.RepeatDirection = RepeatDirection.Vertical
        rRBL.SelectedIndex = 0
        rRBL.AutoPostBack = True
        AddHandler rRBL.SelectedIndexChanged, AddressOf rRBL_SelectedIndexChanged
        tCell.Controls.Add(rRBL)
        tRow.Controls.Add(tCell)
        FT.Rows.Add(tRow)
    End Sub
    Protected Sub rRBL_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)
        Response.Redirect("~/signoff/signoff.aspx?smid=sg&smode=1&formtypeindex=" & DDLFormType.SelectedIndex &
          "&formstatusindex=" & DDLFormStatus.SelectedIndex & "&signflowmode=" & CType(FT.FindControl("signflowmode"), RadioButtonList).SelectedIndex)
    End Sub
    Sub FormListDisplay(displaymode As Integer)
        Dim sfid, formstatus As Integer
        Dim str() As String
        'Dim match As Boolean

        sfid = 0
        'formstatus = 0
        ds.Reset()
        SetGridViewStyle()
        SetFormListGridViewFields()
        If (DDLFormType.SelectedIndex <> 0) Then
            str = Split(DDLFormType.SelectedValue, " ")
            sfid = str(1)
        End If
        If (DDLFormStatus.SelectedIndex = 0) Then '關卡
            formstatus = 1
        ElseIf (DDLFormStatus.SelectedIndex = 1) Then '送審
            formstatus = 100
        ElseIf (DDLFormStatus.SelectedIndex = 2) Then '已核准
            formstatus = 2
        ElseIf (DDLFormStatus.SelectedIndex = 3) Then '已駁回
            formstatus = 3
        ElseIf (DDLFormStatus.SelectedIndex = 4) Then '我的待送審單
            formstatus = 101
        ElseIf (DDLFormStatus.SelectedIndex = 5) Then '作廢
            formstatus = 5
        ElseIf (DDLFormStatus.SelectedIndex = 6) Then '待歸檔
            formstatus = 102
        ElseIf (DDLFormStatus.SelectedIndex = 7) Then '已歸檔
            formstatus = 103
        ElseIf (DDLFormStatus.SelectedIndex = 8) Then '待知悉
            formstatus = 105
        ElseIf (DDLFormStatus.SelectedIndex = 9) Then '已知悉
            formstatus = 104
        End If
        'Dim count As Integer
        Dim sfid_rule As String
        If (sfid = 0) Then
            sfid_rule = ""
        Else
            sfid_rule = " and T1.sfid=" & sfid
        End If

        If (formstatus = 100) Then '已送審 
            SqlCmd = "SELECT T1.price,T1.priceunit,T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate " &
                 "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum " &
                 "where T0.uid='" & sid & "' and T0.status=100" & sfid_rule & 'and T1.sfid=" & sfid &
                " order by T1.sfid,T1.docnum desc"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1)
            connsap1.Close()
        ElseIf (formstatus = 102) Then '待歸檔
            SqlCmd = "Select T1.price,T1.priceunit,T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate,T0.seq " &
                 "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum " &
                 "where T0.signprop=1 And T0.status=1" & sfid_rule & " and T0.uid ='" & sid & "' " &
                 " order by T1.sfid,T1.docnum desc"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1)
            connsap1.Close()

        ElseIf (formstatus = 103) Then '已歸檔或只用@XASCH T1.sid=sid and T1.status='T'... 
            SqlCmd = "Select T1.price,T1.priceunit,T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate,T0.seq " &
                 "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum " &
                 "where T0.signprop=1 And T1.status='T'" & sfid_rule & " and T0.uid ='" & sid & "' " &
                 " order by T1.sfid,T1.docnum desc"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1)
            connsap1.Close()

        ElseIf (formstatus = 101) Then '我的待送審單
            SqlCmd = "SELECT T1.price,T1.priceunit,T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate " &
                         "FROM dbo.[@XASCH] T1 " &
                         "where T1.sid='" & sid & "' and (T1.status='E' or T1.status='D' or T1.status='B' or T1.status='R')" & sfid_rule &
                         " order by T1.sfid,T1.docnum desc"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1)
            connsap1.Close()
        ElseIf (formstatus = 1) Then '待簽核(包含未送審 , 關卡簽核 , 待歸檔,待知悉)
            SqlCmd = "SELECT T1.price,T1.priceunit,T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate " &
                         "FROM dbo.[@XASCH] T1 " &
                         "where T1.sid='" & sid & "' and (T1.status='E' or T1.status='D' or T1.status='B' or T1.status='R')" & sfid_rule &
                         " order by T1.sfid,T1.docnum desc"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1) '未送審
            connsap1.Close()
            SqlCmd = "Select T1.price,T1.priceunit,T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate,T0.seq " &
                 "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum " &
                 "where T0.signprop=0 and T0.status=1 and T1.status<>'B' and T1.status<>'R'" & sfid_rule & " and T0.uid='" & sid & "' " &
                 " order by T0.signprop,T1.sfid,T1.docnum desc"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1) '關卡簽核
            connsap1.Close()
            SqlCmd = "Select T1.price,T1.priceunit,T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate,T0.seq " &
                 "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum " &
                 "where T0.signprop=1 And T1.status='F'" & sfid_rule & " and T0.uid='" & sid & "' " &
                 " order by T0.signprop,T1.sfid,T1.docnum desc"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1) '待歸檔
            connsap1.Close()
            SqlCmd = "Select T1.price,T1.priceunit,T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate,T0.seq " &
                 "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum " &
                 "where T0.signprop=2 And T0.status=1" & sfid_rule & " and T0.uid ='" & sid & "' " &
                 " order by T1.sfid,T1.docnum desc"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1) '待知悉
            connsap1.Close()
        ElseIf (formstatus = 104) Then '已知悉
            SqlCmd = "Select T1.price,T1.priceunit,T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate,T0.seq " &
                 "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum " &
                 "where T0.signprop=2 And T0.status=104" & sfid_rule & " and T0.uid ='" & sid & "' " &
                 " order by T1.sfid,T1.docnum desc"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1)
            connsap1.Close()
        ElseIf (formstatus = 105) Then '待知悉
            SqlCmd = "Select T1.price,T1.priceunit,T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate,T0.seq " &
                 "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum " &
                 "where T0.signprop=2 And T0.status=1" & sfid_rule & " and T0.uid ='" & sid & "' " &
                 " order by T1.sfid,T1.docnum desc"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1)
            connsap1.Close()
        Else
            SqlCmd = "SELECT T1.price,T1.priceunit,T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate,T0.seq " &
                 "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum " &
                 "where T0.status=" & formstatus & " " & sfid_rule & " And T0.uid ='" & sid & "' " &
                " order by T1.sfid,T1.docnum desc"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1)
            connsap1.Close()
        End If

        ds.Tables(0).Columns.Add("action")
        ds.Tables(0).Columns.Add("recall")
        ds.Tables(0).Columns.Add("addsignoff")
        'If (ds.Tables(0).Rows.Count = 0) Then
        '    CommUtil.ShowMsg(Me, "無任何資料")
        'End If
        gv1.DataSource = ds.Tables(0)
        gv1.DataBind()

    End Sub
    Sub SetGridViewStyle()
        gv1.AutoGenerateColumns = False
        'gv1.ShowHeader = True
        gv1.AllowPaging = True
        gv1.PageSize = 20
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

    Sub SetFormListGridViewFields()
        Dim oBoundField As BoundField
        gv1.Columns.Clear()
        oBoundField = New BoundField
        oBoundField.HeaderText = "單號"
        oBoundField.DataField = "docnum"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "表單種類"
        oBoundField.DataField = "sfid"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "主旨"
        oBoundField.DataField = "subject"
        oBoundField.ShowHeader = True
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "金額"
        oBoundField.DataField = "price"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:N0}"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "幣別"
        oBoundField.DataField = "priceunit"
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ShowHeader = True
        oBoundField.HtmlEncode = False
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "製單人"
        oBoundField.DataField = "issuedperson"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "狀態"
        oBoundField.DataField = "status"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "可抽回"
        oBoundField.DataField = "recall"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "日期"
        oBoundField.DataField = "docdate"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "動作"
        oBoundField.DataField = "action"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "增簽"
        oBoundField.DataField = "addsignoff"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)
    End Sub
    Sub SearchDoc()
        'Dim drL As SqlDataReader
        Dim connL As New SqlConnection
        ds.Reset()
        SetGridViewStyle()
        SetFormListGridViewFields()
        If (TxtDocnum.Text <> "") Then
            SqlCmd = "SELECT T1.price,T1.priceunit,T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate,seq=1,T1.sname As uname,T1.sid As uid " &
             "FROM dbo.[@XASCH] T1 " &
             "where (status='E' or status='D') and docnum=" & TxtDocnum.Text & " and sid='" & sid & "'" &
             " order by T1.sfid,T1.docnum desc"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1)
            connsap1.Close()
            SqlCmd = "Select distinct T1.price,T1.priceunit,T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate " &
                    "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum " &
                    "where T1.status<>'E' and T1.status<>'D' and docentry=" & TxtDocnum.Text & " and uid='" & sid & "'" &
                    " order by T1.sfid,T1.docnum desc"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1)
            connsap1.Close()

            'SqlCmd = "Select count(*) " &
            '        "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum " &
            '        "where T1.status<>'E' and T1.status<>'D' and docentry=" & TxtDocnum.Text & " and uid='" & sid & "'"
            'drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
            'drL.Read()
            'If (drL(0) = 1) Then
            '    SqlCmd = "Select T1.price,T1.priceunit,T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate,T0.seq " &
            '        "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum " &
            '        "where T1.status<>'E' and T1.status<>'D' and docentry=" & TxtDocnum.Text & " and uid='" & sid & "'" &
            '        " order by T1.sfid,T1.docnum desc"
            '    ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1)
            '    connsap1.Close()
            'ElseIf (drL(0) > 1) Then
            '    SqlCmd = "Select T1.price,T1.priceunit,T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate,T0.seq " &
            '        "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum " &
            '        "where T0.signprop<>1 and T1.status<>'E' and T1.status<>'D' and docentry=" & TxtDocnum.Text & " and uid='" & sid & "'" &
            '        " order by T1.sfid,T1.docnum desc"
            '    ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1)
            '    connsap1.Close()
            'End If
            'drL.Close()
            'connL.Close()
        ElseIf (TxtKW.Text <> "") Then
            SqlCmd = "SELECT T1.price,T1.priceunit,T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate,seq=1,T1.sname As uname,T1.sid As uid " &
             "FROM dbo.[@XASCH] T1 " &
             "where (status='E' or status='D') and sid='" & sid & "' and subject like '%" & TxtKW.Text & "%' order by T1.sfid,T1.docnum desc"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1)
            connsap1.Close()
            SqlCmd = "Select distinct T1.price,T1.priceunit,T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate " &
                    "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum " &
                    "where uid='" & sid & "' and subject like '%" & TxtKW.Text & "%' order by T1.sfid,T1.docnum desc"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1)
            connsap1.Close()
        Else
            CommUtil.ShowMsg(Me, "表單號或主旨關鍵字須輸入")
            Exit Sub
        End If

        ds.Tables(0).Columns.Add("action")
        ds.Tables(0).Columns.Add("recall")
        ds.Tables(0).Columns.Add("addsignoff")
        If (ds.Tables(0).Rows.Count = 0) Then
            CommUtil.ShowMsg(Me, "查無條件設定之所屬表單")
        End If
        gv1.DataSource = ds.Tables(0)
        gv1.DataBind()
    End Sub
End Class