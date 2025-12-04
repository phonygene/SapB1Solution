Imports System.Data
Imports System.Data.SqlClient
Imports AjaxControlToolkit
'Imports System.Net.Mail
'Imports System.Web.Mail
Public Class signoffsetup
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public connsap, conn, conndept As New SqlConnection
    Public SqlCmd As String
    Public dr, dr1, drdept As SqlDataReader
    Public ds As New DataSet
    Public DDLFormType, DDLSignDefault As DropDownList
    Public BtnSave As Button
    Public ScriptManager1 As New ScriptManager
    Public sfid As Integer
    Public mode As String
    Public RBLSignMethod As RadioButtonList
    Public ChkDel As CheckBox
    Public TxtSignPName As TextBox
    Public LabelDefault As Label
    Public signpersonmaxrow As Integer
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        Page.Form.Controls.Add(ScriptManager1)
        Dim perm As String
        signpersonmaxrow = 17
        FTCreate()
        perm = CommUtil.GetAssignRight("sg300", Session("s_id"))
        If (InStr(perm, "m") Or InStr(perm, "n") Or InStr(perm, "d")) Then
            'RBLSignMethod.Enabled = True
        Else
            RBLSignMethod.Enabled = False
            RBLSignMethod.SelectedIndex = 1
            DDLFormType.Visible = False
            DDLSignDefault.Visible = True
            LabelDefault.Visible = True
            TxtSignPName.Visible = True
        End If
        CreateSignFlowPerson()
        If (IsPostBack) Then
            sfid = ViewState("sfid")
            'CreateSignFlowPerson()
        Else
            mode = Request.QueryString("mode")
            DDLFormType.SelectedIndex = Request.QueryString("formstatusindex")
            DDLSignDefault.SelectedIndex = Request.QueryString("ddlsigndefaultindex")
            sfid = Request.QueryString("sfid")
            ViewState("sfid") = sfid
            'If (mode = "init") Then

            'End If
            If (mode = "select") Then
                'CreateSignFlowPerson()
                PutDataToForm()
            End If

            If (mode = "ctinit") Then
                DDLFormType.Visible = False
                DDLSignDefault.Visible = True
                CT.Visible = True
                DDLSignDefault.SelectedIndex = 0
                RBLSignMethod.SelectedIndex = 1
                'CreateSignFlowPerson()
                LabelDefault.Visible = True
                TxtSignPName.Visible = True
                ChkDel.Visible = False
                'PutDataToForm()
            End If
            If (mode = "ctimport" Or mode = "saveokdefault" Or mode = "deldefault") Then
                DDLFormType.Visible = False
                DDLSignDefault.Visible = True
                CT.Visible = True
                RBLSignMethod.SelectedIndex = 1
                'CreateSignFlowPerson()
                If (mode = "deldefault") Then
                    DeleteMemberOfSingOffFlow()
                End If
                PutDefaultDataToForm()
                LabelDefault.Visible = False
                TxtSignPName.Visible = False
                ChkDel.Visible = True
            End If
            If (mode = "del") Then '要放這理 , CreateSignFlowPerson()要在之前 , 不然有些物件尚未建立
                DeleteMemberOfSingOffFlow()
            End If
            If (mode = "saveok") Then
                If (Request.QueryString("fieldcheckwarning") = 0) Then
                    CommUtil.ShowMsg(Me, "內定簽核群組更新完成")
                Else
                    CommUtil.ShowMsg(Me, "簽核表單設定歸檔屬性有" & Request.QueryString("propc") & "個,但發現有" & Request.QueryString("fieldcheckwarning") & "個未設定部門排除,請設定")
                End If
                PutDataToForm()
            ElseIf (mode = "saveokdefault") Then
                If (Request.QueryString("fieldcheckwarning") = 0) Then
                    CommUtil.ShowMsg(Me, "預設簽核群組更新完成")
                Else
                    CommUtil.ShowMsg(Me, "簽核表單設定歸檔屬性有" & Request.QueryString("propc") & "個,但發現有" & Request.QueryString("fieldcheckwarning") & "個未設定部門排除,請設定")
                End If
            End If
        End If
    End Sub
    Sub FTCreate()
        Dim tCell As TableCell
        Dim tRow As TableRow
        tRow = New TableRow()
        tCell = New TableCell()
        RBLSignMethod = New RadioButtonList()
        RBLSignMethod.Items.Add("內定表單之簽核人員")
        RBLSignMethod.Items.Add("自訂表單之個人預設簽核人員")
        RBLSignMethod.Items(0).Value = 0
        RBLSignMethod.Items(1).Value = 1
        RBLSignMethod.ID = "signmethod_rblist"
        RBLSignMethod.RepeatDirection = RepeatDirection.Vertical
        RBLSignMethod.SelectedIndex = 0
        RBLSignMethod.AutoPostBack = True
        AddHandler RBLSignMethod.SelectedIndexChanged, AddressOf RBLSignMethod_SelectedIndexChanged
        tCell.Controls.Add(RBLSignMethod)
        tRow.Cells.Add(tCell)

        tCell = New TableCell()
        'tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Left
        DDLFormType = New DropDownList
        DDLFormType.ID = "ddl_formtype"
        DDLFormType.Width = 150
        DDLFormType.AutoPostBack = True
        SqlCmd = "Select T0.sfname,T0.sfid from dbo.[@XSFTT] T0 where sftype<>3 order by sfid"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            DDLFormType.Items.Clear()
            DDLFormType.Items.Add("請選擇表單設置")
            Do While (dr.Read())
                DDLFormType.Items.Add(dr(0) & " " & dr(1))
            Loop
        End If
        dr.Close()
        connsap.Close()
        AddHandler DDLFormType.SelectedIndexChanged, AddressOf DDLFormType_SelectedIndexChanged
        tCell.Controls.Add(DDLFormType)

        DDLSignDefault = New DropDownList
        DDLSignDefault.ID = "ddl_signdefault"
        DDLSignDefault.Width = 150
        DDLSignDefault.AutoPostBack = True
        SqlCmd = "Select distinct T0.signpid,T0.signpname from dbo.[@XSPAT] T0 where ownid='" & Session("s_id") & "'"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        DDLSignDefault.Items.Clear()
        DDLSignDefault.Items.Add("新增預設簽核群組")
        If (dr.HasRows) Then
            Do While (dr.Read())
                DDLSignDefault.Items.Add(dr(1) & " " & dr(0))
            Loop
        End If
        dr.Close()
        connsap.Close()
        AddHandler DDLSignDefault.SelectedIndexChanged, AddressOf DDLSignDefault_SelectedIndexChanged
        DDLSignDefault.Visible = False
        tCell.Controls.Add(DDLSignDefault)
        tRow.Cells.Add(tCell)
        LabelDefault = New Label()
        LabelDefault.ID = "label_ddlsigndefault1"
        LabelDefault.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp預設簽核名稱:"
        tCell.Controls.Add(LabelDefault)
        TxtSignPName = New TextBox
        TxtSignPName.ID = "txt_signpname"
        TxtSignPName.Visible = False
        tCell.Controls.Add(TxtSignPName)
        tRow.Controls.Add(tCell)
        TxtSignPName.Visible = False
        LabelDefault.Visible = False
        FT.Rows.Add(tRow)
    End Sub
    Protected Sub RBLSignMethod_SelectedIndexChanged(sender As Object, e As EventArgs)
        If (RBLSignMethod.SelectedIndex = 0) Then
            DDLFormType.Visible = True
            DDLSignDefault.Visible = False
            CT.Visible = False
            DDLFormType.SelectedIndex = 0
            TxtSignPName.Visible = False
            LabelDefault.Visible = False
        Else
            'DDLFormType.Visible = False
            'DDLSignDefault.Visible = True
            'CT.Visible = True
            'DDLSignDefault.SelectedIndex = 0
            'ClearCTItems()
            '以上mark起來部份 , 也可以如此用 , 但也可如下述重新refresh , 但上述mark起來部份要移至Page_load中
            Response.Redirect("~/signoff/signoffsetup.aspx?smid=sg&smode=3&mode=ctinit")
        End If
    End Sub
    Protected Sub DDLSignDefault_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim sfid As Integer
        If (DDLSignDefault.SelectedIndex = 0) Then
            Response.Redirect("~/signoff/signoffsetup.aspx?smid=sg&smode=3&mode=ctinit")
        Else
            Dim str() As String
            str = Split(DDLSignDefault.SelectedValue, " ")
            sfid = str(1)
            'CT.Visible = True
            Response.Redirect("~/signoff/signoffsetup.aspx?smid=sg&smode=3&mode=ctimport&ddlsigndefaultindex=" & DDLSignDefault.SelectedIndex & "&sfid=" & sfid)
            'PutDataToForm()
        End If
    End Sub
    Sub CreateSignFlowPerson()
        Dim i As Integer
        Dim tRow As TableRow
        Dim tCell As TableCell
        Dim RBLProp, RBLSkip As RadioButtonList
        Dim TxtID As TextBox
        'Dim BtnDel As Button
        Dim idstr As String
        'Dim DDLUser As DropDownExtender
        Dim DDLUser As DropDownList
        Dim Hyper As HyperLink
        Dim Labelx As Label
        tRow = New TableRow()
        tRow.BackColor = Drawing.Color.LightGreen
        tRow.Font.Bold = True
        For i = 1 To 9
            tCell = New TableCell
            tCell.BorderWidth = 1
            tCell.Width = 40
            tCell.HorizontalAlign = HorizontalAlign.Center
            If (i = 1) Then
                tCell.Text = "ID"
                tCell.Width = 80
            ElseIf (i = 2) Then
                tCell.Text = "姓名"
            ElseIf (i = 3) Then
                tCell.Text = "職稱"
            ElseIf (i = 4) Then
                tCell.Text = "屬性"
                tCell.Width = 120
            ElseIf (i = 5) Then
                tCell.Text = "順序"
            ElseIf (i = 6) Then
                tCell.Text = "動作"
            ElseIf (i = 7) Then
                tCell.Text = "記錄號"
                tCell.Width = 50
                tCell.Visible = False
            ElseIf (i = 8) Then
                tCell.Text = "部門排除"
                tCell.Width = 65
            ElseIf (i = 9) Then
                tCell.Text = "簽核進行"
                tCell.Width = 120
            End If
            tRow.Controls.Add(tCell)
        Next
        CT.Rows.Add(tRow)
        For i = 1 To signpersonmaxrow + 1
            tRow = New TableRow()
            tRow.BorderWidth = 1
            tCell = New TableCell
            tCell.BorderWidth = 1
            'LBx = New ListBox
            'LBx.ID = "lb_id_" & i
            'LBx.AutoPostBack = True
            'LBx.Rows = 30
            'AddHandler LBx.SelectedIndexChanged, AddressOf LB_SelectedIndexChanged
            'tCell.Controls.Add(LBx)
            'TxtID = New TextBox
            'TxtID.ID = "txt_id_" & i
            'TxtID.Width = 40
            'tCell.Controls.Add(TxtID)
            'DDLUser = New DropDownExtender
            'DDLUser.TargetControlID = TxtID.ID
            'DDLUser.ID = "ddl_user_" & i
            'DDLUser.DropDownControlID = LBx.ID
            'tCell.Controls.Add(DDLUser)
            'LBx.Items.Clear()
            'LBx.Items.Add("")
            DDLUser = New DropDownList
            DDLUser.ID = "ddl_user_" & i
            DDLUser.Width = 120
            AddHandler DDLUser.SelectedIndexChanged, AddressOf DDLUser_SelectedIndexChanged
            DDLUser.AutoPostBack = True
            tCell.Controls.Add(DDLUser)
            DDLUser.Items.Clear()
            DDLUser.Items.Add("")
            If (RBLSignMethod.SelectedIndex = 0) Then
                'MsgBox("0")
                If (sfid > 50 And sfid < 80) Then
                    SqlCmd = "select id,name,position from dbo.[user] where email<>'' and topsignoffs=1 and denyf<>1 order by branch,grp"
                Else
                    SqlCmd = "select id,name,position from dbo.[user] where email<>'' and denyf<>1 order by branch,grp"
                End If
            Else
                'MsgBox("1")
                SqlCmd = "select id,name,position from dbo.[user] where denyf<>1 and email<>'' and id <>'" & Session("s_id") & "' order by branch,grp"
            End If
            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
            If (dr.HasRows) Then
                Do While (dr.Read())
                    idstr = dr(0) & " " & dr(1) & " " & dr(2)
                    DDLUser.Items.Add(idstr)
                Loop
            End If
            dr.Close()
            conn.Close()
            tRow.Controls.Add(tCell)

            tCell = New TableCell '姓名
            tCell.BorderWidth = 1
            tCell.Wrap = False
            tCell.HorizontalAlign = HorizontalAlign.Center
            tCell.Width = 40
            tRow.Controls.Add(tCell)

            tCell = New TableCell '職稱
            tCell.BorderWidth = 1
            tCell.Wrap = False
            tCell.HorizontalAlign = HorizontalAlign.Center
            tCell.Width = 40
            tRow.Controls.Add(tCell)

            tCell = New TableCell '屬性
            tCell.BorderWidth = 1
            RBLProp = New RadioButtonList
            RBLProp.ID = "rbl_prop_" & i
            RBLProp.Width = 180
            RBLProp.RepeatDirection = RepeatDirection.Vertical
            RBLProp.Items.Add("簽核")
            RBLProp.Items.Add("歸檔")
            RBLProp.Items.Add("知悉")
            RBLProp.Items(0).Value = 0
            RBLProp.Items(1).Value = 1
            RBLProp.Items(2).Value = 2
            RBLProp.AutoPostBack = True
            RBLProp.Enabled = False
            AddHandler RBLProp.SelectedIndexChanged, AddressOf RBLProp_SelectedIndexChanged
            tCell.Controls.Add(RBLProp)
            tRow.Controls.Add(tCell)

            tCell = New TableCell '順序
            tCell.BorderWidth = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            TxtID = New TextBox
            TxtID.ID = "txt_seq_" & i
            TxtID.Width = 40
            tCell.Controls.Add(TxtID)
            tRow.Controls.Add(tCell)

            tCell = New TableCell '刪除
            tCell.BorderWidth = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            Hyper = New HyperLink
            Hyper.ID = "hyper_del_" & i
            Hyper.Text = "刪除"
            Hyper.Enabled = False
            tCell.Controls.Add(Hyper)
            tRow.Controls.Add(tCell)

            tCell = New TableCell '記錄號
            tCell.BorderWidth = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            tCell.Visible = False
            tRow.Controls.Add(tCell)

            tCell = New TableCell '
            tCell.BorderWidth = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            Hyper = New HyperLink
            Hyper.ID = "hyper_dptexclude_" & i
            Hyper.Text = "無"
            Hyper.Enabled = False
            tCell.Controls.Add(Hyper)
            tRow.Controls.Add(tCell)

            tCell = New TableCell 'Skip
            tCell.BorderWidth = 1
            RBLSkip = New RadioButtonList
            RBLSkip.ID = "rbl_skip_" & i
            RBLSkip.Width = 150
            RBLSkip.RepeatDirection = RepeatDirection.Vertical
            RBLSkip.Items.Add("正常")
            RBLSkip.Items.Add("暫時Skip")
            RBLSkip.Items(0).Value = 0
            RBLSkip.Items(1).Value = 1
            'RBLSkip.AutoPostBack = True
            'RBLSkip.Enabled = False
            'AddHandler RBLSkip.SelectedIndexChanged, AddressOf RBLSkip_SelectedIndexChanged
            tCell.Controls.Add(RBLSkip) 'ooo
            tRow.Controls.Add(tCell)

            CT.Rows.Add(tRow)
        Next
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.BackColor = Drawing.Color.LightGreen
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 8
        tCell.HorizontalAlign = HorizontalAlign.Center
        BtnSave = New Button
        BtnSave.ID = "btn_save_" & i
        BtnSave.Width = 40
        BtnSave.Text = "儲存"
        BtnSave.OnClientClick = "Return confirm('要存檔嗎')"
        AddHandler BtnSave.Click, AddressOf BtnSave_Click
        tCell.Controls.Add(BtnSave)
        Labelx = New Label()
        Labelx.ID = "label_upfile"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        ChkDel = New CheckBox
        ChkDel.ID = "chk_del"
        ChkDel.Text = "刪除檔案"
        ChkDel.AutoPostBack = True
        If (RBLSignMethod.SelectedIndex = 1 And DDLSignDefault.SelectedIndex <> 0) Then
            ChkDel.Visible = True
        Else
            ChkDel.Visible = False
        End If
        AddHandler ChkDel.CheckedChanged, AddressOf ChkDel_CheckedChanged
        tCell.Controls.Add(ChkDel)

        tRow.Controls.Add(tCell)
        CT.Rows.Add(tRow)
    End Sub
    Protected Sub ChkDel_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)

        If (ChkDel.Checked) Then
            BtnSave.Text = "刪除"
        Else
            BtnSave.Text = "儲存"
        End If
    End Sub
    Sub PutDataToForm()
        Dim str() As String
        Dim i As Integer
        Dim seq As Integer
        i = 1
        seq = 1
        str = Split(DDLFormType.SelectedValue, " ")
        sfid = str(1)
        SqlCmd = "select T0.uid,T0.seq,T0.prop,T0.num,T0.process from [@XSPMT] T0 where T0.sfid=" & sfid & " order by T0.seq"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            Do While (dr.Read())
                SqlCmd = "select T0.name,T0.position from dbo.[User] T0 where T0.id='" & dr(0) & "'"
                dr1 = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                dr1.Read()
                CType(CT.FindControl("ddl_user_" & i), DropDownList).SelectedValue = dr(0) & " " & dr1(0) & " " & dr1(1)
                CT.Rows(i).Cells(1).Text = dr1(0)
                CT.Rows(i).Cells(2).Text = dr1(1)
                CType(CT.FindControl("rbl_prop_" & i), RadioButtonList).SelectedValue = dr(2)
                CType(CT.FindControl("rbl_prop_" & i), RadioButtonList).Enabled = True
                CType(CT.FindControl("txt_seq_" & i), TextBox).Text = dr(1)
                If (dr(1) <> 0) Then
                    CType(CT.FindControl("txt_seq_" & i), TextBox).Enabled = True
                Else
                    CType(CT.FindControl("txt_seq_" & i), TextBox).Enabled = False
                End If
                CType(CT.FindControl("hyper_del_" & i), HyperLink).NavigateUrl = "~/signoff/signoffsetup.aspx?smid=sg&smode=3&num=" & dr(3) & "&mode=del&formstatusindex=" & DDLFormType.SelectedIndex & "&row=" & i
                CType(CT.FindControl("hyper_del_" & i), HyperLink).Enabled = True
                CT.Rows(i).Cells(6).Text = dr(3)
                CType(CT.FindControl("hyper_dptexclude_" & i), HyperLink).NavigateUrl = "~/signoff/signflowexclude.aspx?smid=sg&smode=3&sfid=" & sfid & "&signid=" & dr(0) & "&formstatusindex=" & DDLFormType.SelectedIndex
                CType(CT.FindControl("hyper_dptexclude_" & i), HyperLink).Enabled = True
                SqlCmd = "select count(*) from [dbo].[@XSDET] where uid='" & dr(0) & "' and sfid=" & sfid
                drdept = CommUtil.SelectSapSqlUsingDr(SqlCmd, conndept)
                drdept.Read()
                If (drdept(0) <> 0) Then
                    CType(CT.FindControl("hyper_dptexclude_" & i), HyperLink).Text = "有"
                    CType(CT.FindControl("ddl_user_" & i), DropDownList).Enabled = False
                    CType(CT.FindControl("rbl_prop_" & i), RadioButtonList).Enabled = False
                Else
                    CType(CT.FindControl("ddl_user_" & i), DropDownList).Enabled = True
                End If
                CType(CT.FindControl("rbl_skip_" & i), RadioButtonList).SelectedValue = dr(4)
                drdept.Close()
                conndept.Close()

                dr1.Close()
                conn.Close()
                i = i + 1
            Loop
        End If
        dr.Close()
        connsap.Close()
    End Sub
    Sub PutDefaultDataToForm()
        Dim i As Integer
        Dim seq As Integer
        Dim first As Boolean
        first = True
        i = 1
        seq = 1
        SqlCmd = "select T0.uid,T0.seq,T0.prop,T0.num,T0.signpname,T0.process from [@XSPAT] T0 where T0.signpid=" & sfid & " order by T0.seq"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            Do While (dr.Read())
                If (first) Then
                    DDLSignDefault.SelectedValue = dr(4) & " " & sfid
                    first = False
                End If
                SqlCmd = "select T0.name,T0.position from dbo.[User] T0 where T0.id='" & dr(0) & "'"
                dr1 = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                dr1.Read()
                CType(CT.FindControl("ddl_user_" & i), DropDownList).SelectedValue = dr(0) & " " & dr1(0) & " " & dr1(1)
                CT.Rows(i).Cells(1).Text = dr1(0)
                CT.Rows(i).Cells(2).Text = dr1(1)
                CType(CT.FindControl("rbl_prop_" & i), RadioButtonList).SelectedValue = dr(2)
                CType(CT.FindControl("rbl_prop_" & i), RadioButtonList).Enabled = True
                CType(CT.FindControl("txt_seq_" & i), TextBox).Text = dr(1)
                If (dr(1) <> 0) Then
                    CType(CT.FindControl("txt_seq_" & i), TextBox).Enabled = True
                Else
                    CType(CT.FindControl("txt_seq_" & i), TextBox).Enabled = False
                End If
                CType(CT.FindControl("hyper_del_" & i), HyperLink).NavigateUrl = "~/signoff/signoffsetup.aspx?smid=sg&smode=3&num=" & dr(3) & "&mode=deldefault&ddlsigndefaultindex=" & DDLSignDefault.SelectedIndex & "&row=" & i & "&sfid=" & sfid
                CType(CT.FindControl("hyper_del_" & i), HyperLink).Enabled = True
                CT.Rows(i).Cells(6).Text = dr(3)
                CType(CT.FindControl("rbl_skip_" & i), RadioButtonList).SelectedValue = dr(5)
                dr1.Close()
                conn.Close()
                i = i + 1
            Loop
        End If
        dr.Close()
        connsap.Close()
    End Sub
    Protected Sub DDLFormType_SelectedIndexChanged(sender As Object, e As EventArgs)
        If (DDLFormType.SelectedIndex = 0) Then
            CT.Visible = False
        Else
            Dim str() As String
            str = Split(DDLFormType.SelectedValue, " ")
            sfid = str(1)
            CT.Visible = True
            Response.Redirect("~/signoff/signoffsetup.aspx?smid=sg&smode=3&mode=select&formstatusindex=" & DDLFormType.SelectedIndex & "&sfid=" & sfid)
            'PutDataToForm()
        End If

    End Sub
    Protected Sub DDLUser_SelectedIndexChanged(sender As Object, e As EventArgs) '123456
        If ((RBLSignMethod.SelectedIndex = 0 And DDLFormType.SelectedIndex <> 0) Or RBLSignMethod.SelectedIndex = 1) Then
            Dim idstr As String
            Dim str() As String
            Dim row As Integer
            idstr = sender.ID
            str = Split(idstr, "_")
            row = CInt(str(2))
            If (sender.SelectedIndex <> 0) Then
                str = Split(sender.SelectedValue, " ")
                'id field
                CType(CT.FindControl("ddl_user_" & row), DropDownList).SelectedValue = str(0) & " " & str(1) & " " & str(2)
                CT.Rows(row).Cells(1).Text = str(1)
                CT.Rows(row).Cells(2).Text = str(2)
                'CType(CT.FindControl("hyper_dptexclude_" & row), HyperLink).NavigateUrl = "~/signoff/signflowexclude.aspx?smid=sg&smode=3" &
                '                    "&sfid=" & sfid & "&signid=" & str(0) & "&formstatusindex=" & DDLFormType.SelectedIndex
                'CType(CT.FindControl("hyper_dptexclude_" & row), HyperLink).Enabled = True
            Else
                CType(CT.FindControl("ddl_user_" & row), DropDownList).SelectedValue = ""
                CT.Rows(row).Cells(1).Text = ""
                CT.Rows(row).Cells(2).Text = ""
            End If
            CType(CT.FindControl("rbl_prop_" & row), RadioButtonList).Enabled = True
            'CT.Rows(row).Cells(4).Text = row
            CType(CT.FindControl("txt_seq_" & row), TextBox).Text = row
            CType(CT.FindControl("rbl_skip_" & row), RadioButtonList).SelectedValue = 0
            ViewState("sfid") = sfid
        Else
            CommUtil.ShowMsg(Me, "須先選擇預設置之表單")
        End If
    End Sub
    Protected Sub RBLProp_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim idstr As String
        Dim str() As String
        Dim row As Integer
        idstr = sender.ID
        str = Split(idstr, "_")
        row = CInt(str(2))
        If (sender.SelectedIndex = 0) Then
            CType(CT.FindControl("txt_seq_" & row), TextBox).Enabled = True
            If (CType(CT.FindControl("txt_seq_" & row), TextBox).Text = "0") Then
                CType(CT.FindControl("txt_seq_" & row), TextBox).Text = ""
            End If
        Else
            CType(CT.FindControl("txt_seq_" & row), TextBox).Enabled = False
            CType(CT.FindControl("txt_seq_" & row), TextBox).Text = "0"
        End If
        ViewState("sfid") = sfid
    End Sub
    Protected Sub BtnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) 'ppp
        Dim i, j, seq, sighcount, cttype, sfid As Integer
        Dim num As Long
        Dim str(), str1() As String
        Dim uid, uname, upos, signpname As String
        Dim prop, propc, targetseq, process As Integer
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim ownid(15) As String
        Dim FieldCheckWarning As Integer
        FieldCheckWarning = 0
        signpname = ""
        ' If (Not ChkDel.Checked) Then
        If (RBLSignMethod.SelectedIndex = 1) Then
            cttype = 1
        Else
            cttype = 0
        End If
        If (cttype = 1 And DDLSignDefault.SelectedIndex = 0) Then
            If (TxtSignPName.Text = "") Then
                CommUtil.ShowMsg(Me, "新增簽核群組沒填名子")
                Exit Sub
            End If
        End If
        propc = 0
        'proprow = 1
        If (cttype = 0) Then
            str = Split(DDLFormType.SelectedValue, " ")
            sfid = str(1)
        Else
            If (DDLSignDefault.SelectedIndex = 0) Then
                SqlCmd = "select IsNull(max(signpid),0) from [dbo].[@XSPAT] "
                dr1 = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                dr1.Read()
                If (dr1(0) <> 0) Then
                    sfid = dr1(0) + 1
                Else
                    sfid = 1
                End If
                dr1.Close()
                connsap.Close()
                signpname = TxtSignPName.Text
            Else
                str = Split(DDLSignDefault.SelectedValue, " ")
                sfid = str(1)
                signpname = str(0)
            End If
        End If
        If (Not ChkDel.Checked) Then
            For i = 1 To signpersonmaxrow + 1
                str1 = Split(CType(CT.FindControl("ddl_user_" & i), DropDownList).SelectedValue, " ")
                uid = str1(0)
                If (uid <> "") Then
                    If (CType(CT.FindControl("rbl_prop_" & i), RadioButtonList).SelectedIndex = -1) Then
                        CommUtil.ShowMsg(Me, uid & "(" & CT.Rows(i).Cells(1).Text & ") 屬性沒設定")
                        Exit Sub
                    End If
                    'MsgBox(CType(CT.FindControl("rbl_prop_" & i), RadioButtonList).SelectedValue)
                    prop = CType(CT.FindControl("rbl_prop_" & i), RadioButtonList).SelectedValue
                    If (prop = 1) Then
                        ownid(propc) = uid
                        propc = propc + 1
                        'proprow = i
                    End If
                    If (Not IsNumeric(CType(CT.FindControl("txt_seq_" & i), TextBox).Text)) Then
                        CommUtil.ShowMsg(Me, uid & "(" & CT.Rows(i).Cells(1).Text & ") 順序沒設定或設定的不是數值")
                        Exit Sub
                    End If
                    targetseq = CInt(CType(CT.FindControl("txt_seq_" & i), TextBox).Text)
                    If (targetseq <> 0) Then
                        For j = i + 1 To signpersonmaxrow + 1
                            If (CType(CT.FindControl("txt_seq_" & j), TextBox).Text <> "") Then
                                If (targetseq = CInt(CType(CT.FindControl("txt_seq_" & j), TextBox).Text)) Then
                                    CommUtil.ShowMsg(Me, "順序" & targetseq & "有重複 , 請修正")
                                    Exit Sub
                                End If
                            End If
                        Next
                    End If
                    For j = i + 1 To signpersonmaxrow + 1
                        If (CType(CT.FindControl("ddl_user_" & j), DropDownList).SelectedValue <> "") Then
                            If (CType(CT.FindControl("ddl_user_" & i), DropDownList).SelectedValue = CType(CT.FindControl("ddl_user_" & j), DropDownList).SelectedValue) Then
                                If (CType(CT.FindControl("rbl_prop_" & i), RadioButtonList).SelectedValue = 0 And CType(CT.FindControl("rbl_prop_" & j), RadioButtonList).SelectedValue = 1) Then
                                    'nothing
                                ElseIf (CType(CT.FindControl("rbl_prop_" & i), RadioButtonList).SelectedValue = 1 And CType(CT.FindControl("rbl_prop_" & j), RadioButtonList).SelectedValue = 0) Then
                                    'nothing
                                Else
                                    CommUtil.ShowMsg(Me, "簽核id-" & uid & "(" & CT.Rows(i).Cells(1).Text & ")有重複 , 請修正")
                                    Exit Sub
                                End If
                            End If
                        End If
                    Next
                End If
            Next
            ownid(propc) = "end"
            Dim k As Integer
            k = 0
            If (propc >= 2) Then
                Do While (ownid(k) <> "end")
                    SqlCmd = "select count(*) from [dbo].[@XSDET] where uid='" & ownid(k) & "' and sfid=" & sfid
                    drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                    If (drL.HasRows) Then
                        drL.Read()
                        If (drL(0) = 0) Then
                            'CommUtil.ShowMsg(Me, "簽核表單歸檔設定有" & propc & "個,需設定部門排除,但id " & ownid(k) & " 無設定排除,請重設")
                            FieldCheckWarning = FieldCheckWarning + 1
                            'Exit Sub
                        End If
                    End If
                    drL.Close()
                    connL.Close()
                    k = k + 1
                Loop
            End If
            sighcount = 0
            For i = 1 To signpersonmaxrow + 1
                str1 = Split(CType(CT.FindControl("ddl_user_" & i), DropDownList).SelectedValue, " ")
                uid = str1(0)
                If (uid <> "") Then
                    seq = CInt(CType(CT.FindControl("txt_seq_" & i), TextBox).Text)
                    uname = CT.Rows(i).Cells(1).Text
                    upos = CT.Rows(i).Cells(2).Text
                    prop = CType(CT.FindControl("rbl_prop_" & i), RadioButtonList).SelectedValue
                    process = CType(CT.FindControl("rbl_skip_" & i), RadioButtonList).SelectedValue
                    If (CT.Rows(i).Cells(6).Text = "") Then
                        If (cttype = 0) Then
                            SqlCmd = "insert into [dbo].[@XSPMT] (sfid,uid,seq,prop,process) " &
                                "values(" & sfid & ",'" & uid & "'," & seq & "," & prop & "," & process & ")"
                        Else
                            SqlCmd = "insert into [dbo].[@XSPAT] (signpid,uid,seq,prop,signpname,ownid,process) " &
                                "values(" & sfid & ",'" & uid & "'," & seq & "," & prop & ",'" & signpname & "','" & Session("s_id") & "'," & process & ")" 'ron
                        End If
                        CommUtil.SqlSapExecute("ins", SqlCmd, connsap)
                        connsap.Close()
                    Else
                        num = CLng(CT.Rows(i).Cells(6).Text)
                        If (cttype = 0) Then
                            SqlCmd = "Update [dbo].[@XSPMT] set uid='" & uid & "',prop='" & prop & "',seq=" & seq & ",process=" & process & " " &
                        "where num=" & num
                        Else
                            SqlCmd = "Update [dbo].[@XSPAT] set uid='" & uid & "',prop='" & prop & "',seq=" & seq & ",process=" & process & " " &
                        "where num=" & num
                        End If
                        CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                        connsap.Close()
                    End If
                End If
            Next
            '以seq排序讀出此設定, 再以1開始,重新順序寫入seq
            If (cttype = 0) Then
                SqlCmd = "select T0.seq,T0.num from [@XSPMT] T0 where T0.prop=0 and T0.sfid=" & sfid & " order by T0.seq"
            Else
                SqlCmd = "select T0.seq,T0.num from [@XSPAT] T0 where T0.prop=0 and T0.signpid=" & sfid & " order by T0.seq"
            End If
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                i = 1
                Do While (dr.Read())
                    If (i <> dr(0)) Then
                        If (cttype = 0) Then
                            SqlCmd = "Update [dbo].[@XSPMT] set seq=" & i & " " &
                        "where num=" & dr(1)
                        Else
                            SqlCmd = "Update [dbo].[@XSPAT] set seq=" & i & " " &
                        "where num=" & dr(1)
                        End If
                        CommUtil.SqlSapExecute("upd", SqlCmd, conn)
                        conn.Close()
                    End If
                    i = i + 1
                Loop
            End If
            dr.Close()
            connsap.Close()
            If (FieldCheckWarning = 0) Then
                If (cttype = 0) Then
                    Response.Redirect("~/signoff/signoffsetup.aspx?smid=sg&smode=3&mode=saveok&fieldcheckwaring=0&propc=" & propc &
                                      "&formstatusindex=" & DDLFormType.SelectedIndex & "&sfid=" & sfid)
                Else
                    Response.Redirect("~/signoff/signoffsetup.aspx?smid=sg&smode=3&mode=saveokdefault&fieldcheckwaring=0&propc=" & propc &
                                      "&ddlsigndefaultindex=" & DDLSignDefault.SelectedIndex & "&sfid=" & sfid)
                End If
            Else
                If (cttype = 0) Then
                    Response.Redirect("~/signoff/signoffsetup.aspx?smid=sg&smode=3&mode=saveok&fieldcheckwarning=" & FieldCheckWarning & "&propc=" & propc &
                                      "&formstatusindex=" & DDLFormType.SelectedIndex & "&sfid=" & sfid)
                Else
                    Response.Redirect("~/signoff/signoffsetup.aspx?smid=sg&smode=3&mode=saveokdefault&fieldcheckwarning=" & FieldCheckWarning & "&propc=" & propc &
                                      "&ddlsigndefaultindex=" & DDLSignDefault.SelectedIndex & "&sfid=" & sfid)
                End If
            End If
        Else 'delete all records belongs to sfid
            SqlCmd = "delete from [dbo].[@XSPAT] where signpid=" & sfid
            CommUtil.SqlSapExecute("del", SqlCmd, connsap)
            connsap.Close()
            Response.Redirect("~/signoff/signoffsetup.aspx?smid=sg&smode=3&mode=ctinit")
        End If
    End Sub

    'Protected Sub BtnDel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Dim i As Integer
    '    Dim num As Long
    '    Dim str() As String
    '    str = Split(sender.ID, "_")
    '    i = str(2)
    '    num = CLng(CT.Rows(i).Cells(6).Text)
    '    SqlCmd = "delete from [dbo].[@XSPMT] where num=" & num
    '    CommUtil.SqlSapExecute("del", SqlCmd, connsap)
    '    connsap.Close()
    '    Response.Redirect("~/signoff/signoffsetup.aspx?smid=sg&smode=3&mode=del&formstatusindex=" & DDLFormType.SelectedIndex)
    'End Sub
    Sub DeleteMemberOfSingOffFlow()
        Dim index As Integer
        Dim num As Long
        Dim i As Integer
        Dim str() As String
        If (RBLSignMethod.SelectedIndex = 0) Then
            index = Request.QueryString("formstatusindex")
            DDLFormType.SelectedIndex = index
        Else
            index = Request.QueryString("ddlsigndefaultindex")
            DDLSignDefault.SelectedIndex = index
            SqlCmd = "select count(*) from [dbo].[@XSPAT] where signpid=" & sfid
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            If (dr(0) = 1) Then
                CommUtil.ShowMsg(Me, "只剩一個不能刪 , 若要刪除,請勾選刪除checkbox 來刪除此群組")
                dr.Close()
                connsap.Close()
                Exit Sub
            End If
            dr.Close()
            connsap.Close()
        End If
        i = Request.QueryString("row")
        num = Request.QueryString("num")
        If (RBLSignMethod.SelectedIndex = 0) Then
            SqlCmd = "delete from [dbo].[@XSPMT] where num=" & num
        Else
            SqlCmd = "delete from [dbo].[@XSPAT] where num=" & num
        End If
        CommUtil.SqlSapExecute("del", SqlCmd, connsap)
        connsap.Close()
        '以下刪除@XSDET 內有設定部門之record
        If (RBLSignMethod.SelectedIndex = 0) Then
            If (CType(CT.FindControl("hyper_dptexclude_" & i), HyperLink).Text = "有") Then
                str = Split(CType(CT.FindControl("ddl_user_" & i), DropDownList).SelectedValue, " ")
                SqlCmd = "delete from dbo.[@XSDET] where sfid=" & sfid & " and uid='" & str(0) & "'"
                If (CommUtil.SqlSapExecute("del", SqlCmd, connsap)) Then

                Else
                    CommUtil.ShowMsg(Me, "刪除部門排除record失敗")
                End If
                connsap.Close()
            End If
            Response.Redirect("~/signoff/signoffsetup.aspx?smid=sg&smode=3&mode=select&formstatusindex=" & index & "&sfid=" & sfid)
        Else
            Response.Redirect("~/signoff/signoffsetup.aspx?smid=sg&smode=3&mode=ctimport&ddlsigndefaultindex=" & index & "&sfid=" & sfid)
        End If
    End Sub
End Class