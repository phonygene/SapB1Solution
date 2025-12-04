Imports System.Data.SqlClient
Public Class signofftool
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public CommSignOff As New CommSignOff
    Public connsap, conn As New SqlConnection
    Public SqlCmd As String
    Public dr, drsap As SqlDataReader

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        If (Not IsPostBack) Then
            DDLItemAdd()
        End If
    End Sub
    Sub SignOffDataAnalysis()
        Dim str(), delid, mes, delidname, ownidname As String
        Dim replaceid As String
        Dim count As Integer
        Dim connL, connL1 As New SqlConnection
        Dim drL, drL1 As SqlDataReader
        Dim strstatus As String
        Dim replaceflag As Boolean
        count = 1
        mes = ""
        strstatus = ""
        ownidname = ""
        replaceflag = False
        replaceid = ""
        If (DDLDelSIgnUser.SelectedIndex = 0) Then
            CommUtil.ShowMsg(Me, "需選擇欲刪除(異動)之簽核人")
            Exit Sub
        Else
            str = Split(DDLDelSIgnUser.SelectedValue, " ")
            delid = str(0)
            delidname = str(1)
        End If
        If (DDLReplaceSignUser.SelectedIndex = 0) Then
            CommUtil.ShowMsg(Me, "需選擇欲替代之簽核人")
            Exit Sub
        ElseIf (DDLReplaceSignUser.SelectedIndex <> 1) Then
            replaceflag = True
            'str = Split(DDLReplaceSignUser.SelectedValue, " ")
            'replaceid = str(0)
        End If
        MesLB.Items.Clear()
        MesLB.Items.Add("以下若有 *********XXXXXXX******* 表示需手動處理 , 請特別注意")
        MesLB.Items.Add("")
        MesLB.Items.Add("以下為欲刪除ID:" & DDLDelSIgnUser.SelectedValue & " 有關的簽核表需處理List:")
        SqlCmd = "select T0.uid,T1.sfname,T0.prop,T0.sfid from [@XSPMT] T0 inner join [@XSFTT] T1 on T0.sfid=T1.sfid where uid='" & delid & "'"
        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (drsap.HasRows) Then
            Do While (drsap.Read())
                If (drsap(2) = 0) Then
                    mes = count & ". " & drsap(0) & " 單據種類號:" & drsap(3) & " 單據名:" & drsap(1) & " 屬性為: 簽核"
                ElseIf (drsap(2) = 1) Then
                    If (replaceflag) Then
                        mes = count & ". " & drsap(0) & " 單據種類號:" & drsap(3) & " 單據名:" & drsap(1) & " 屬性為: 歸檔"
                    Else
                        mes = count & ". *******" & drsap(0) & " 單據種類號:" & drsap(3) & " 單據名:" & drsap(1) & " 屬性為: 歸檔 *******"
                    End If
                ElseIf (drsap(2) = 2) Then
                    mes = count & ". " & drsap(0) & " 單據種類號:" & drsap(3) & " 單據名:" & drsap(1) & " 屬性為: 知悉"
                End If
                MesLB.Items.Add(mes)
                count = count + 1
            Loop
        Else
            mes = "無任何需更改設定單據存在"
            MesLB.Items.Add(mes)
        End If
        drsap.Close()
        connsap.Close()

        count = 1
        MesLB.Items.Add("")
        MesLB.Items.Add("以下為欲刪除ID:" & DDLDelSIgnUser.SelectedValue & " 有關的自定簽核預設組別將需更改:")
        SqlCmd = "select T0.ownid,T0.prop,T0.num,T0.signpname,T0.signpid from [@XSPAT] T0 where T0.sgtype=0 and T0.uid='" & delid & "' " &
                    "order by T0.ownid"
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            Do While (drL.Read())
                SqlCmd = "select T0.name,T0.position,T0.denyf from dbo.[User] T0 where T0.id='" & drL(0) & "'"
                drL1 = CommUtil.SelectLocalSqlUsingDr(SqlCmd, connL1)
                drL1.Read()
                ownidname = drL1(0)
                drL1.Close()
                connL1.Close()
                If (drL(1) = 0) Then
                    mes = count & ". 自訂簽核人預設組擁有人id:" & drL(0) & " 擁有人名:" & ownidname & " 簽核預設組名" & drL(3) & " 代號:" & drL(4) & " 屬性為: 簽核"
                ElseIf (drL(1) = 1) Then
                    If (replaceflag) Then
                        mes = count & ". 自訂簽核人預設組擁有人id:" & drL(0) & " 擁有人名:" & ownidname & " 簽核預設組名" & drL(3) & " 代號:" & drL(4) & " 屬性為: 歸檔"
                    Else
                        mes = count & ". ******* 自訂簽核人預設組擁有人id:" & drL(0) & " 擁有人名:" & ownidname & " 簽核預設組名" & drL(3) & " 代號:" & drL(4) & " 屬性為: 歸檔 *******"
                    End If
                ElseIf (drL(1) = 2) Then
                    mes = count & ". 自訂簽核人預設組擁有人id:" & drL(0) & " 擁有人名:" & ownidname & " 簽核預設組名" & drL(3) & " 代號:" & drL(4) & " 屬性為: 知悉"
                End If
                MesLB.Items.Add(mes)
                count = count + 1
            Loop
        Else
            mes = "無任何需更改設定單據存在"
            MesLB.Items.Add(mes)
        End If
        drL.Close()
        connL.Close()

        count = 1
        mes = ""
        MesLB.Items.Add(mes)
        MesLB.Items.Add("以下為欲刪除ID:" & DDLDelSIgnUser.SelectedValue & " 有關的需覆核之簽核表單List將需更改:")
        SqlCmd = "select T0.uname,T0.docentry,T1.subject,T2.sfid,T2.sfname,T0.signprop,T0.status,T0.seq from [@XSPWT] T0 Inner Join [@XASCH] T1 On T0.docentry=T1.docnum Inner Join [@XSFTT] T2 On T1.sfid=T2.sfid " &
                "where T0.uid='" & delid & "' and (T0.status = 0 or T0.status = 1 or T0.status=105)"
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            Do While (drL.Read())
                If (drL(6) = 1 And drL(7) = 1) Then
                    If (replaceflag) Then
                        strstatus = "再送審"
                    Else
                        strstatus = "******* 再送審 *******"
                    End If
                ElseIf (drL(6) = 1 And drL(5) = 1) Then
                    If (replaceflag) Then
                        strstatus = "待歸檔"
                    Else
                        strstatus = "******* 待歸檔 *******"
                    End If
                ElseIf (drL(6) = 105) Then
                    strstatus = "待知悉"
                ElseIf (drL(6) = 1) Then
                    strstatus = "待覆核"
                ElseIf (drL(6) = 0 And drL(5) = 1) Then
                    If (replaceflag) Then
                        strstatus = "未到的待歸檔"
                    Else
                        strstatus = "******* 未到的待歸檔 *******"
                    End If
                ElseIf (drL(6) = 0 And drL(5) = 2) Then
                    strstatus = "未到的待知悉"
                ElseIf (drL(6) = 0) Then
                    strstatus = "未到的待覆核"
                End If
                mes = count & ". " & drL(0) & " 簽單號:" & drL(1) & " 狀態:" & strstatus & " 單據種類號:" & drL(3) & " 單據名:" & drL(4) & " 主旨:" & drL(2)
                count = count + 1
                MesLB.Items.Add(mes)
            Loop
        End If
        drL.Close()
        connL.Close()
        If (count = 1) Then
            MesLB.Items.Add("無")
        End If
        '************************
        mes = ""
        MesLB.Items.Add(mes)
        MesLB.Items.Add("以下為欲刪除ID:" & DDLDelSIgnUser.SelectedValue & " 有關的還未送審的簽核表單List將需刪除並告知新繼任者或其他處理方式:")
        count = 1
        SqlCmd = "SELECT T1.docnum,T1.subject,T1.sname,T1.sfid,T1.status,T2.sfname " &
                         "FROM dbo.[@XASCH] T1 Inner Join [@XSFTT] T2 On T1.sfid=T2.sfid " &
                         "where T1.sid='" & delid & "' and (T1.status='E' or T1.status='D') " &
                         " order by T1.sfid,T1.docnum desc"
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            Do While (drL.Read())
                If (drL(4) = "D") Then
                    strstatus = "待送審"
                ElseIf (drL(4) = "E") Then
                    strstatus = "待編輯"
                End If
                If (replaceflag = False) Then
                    strstatus = "******* " & strstatus & " *******"
                End If
                mes = count & ". " & drL(2) & " 簽單號:" & drL(0) & " 狀態:" & strstatus & " 單據種類號:" & drL(3) & " 單據名:" & drL(5) & " 主旨:" & drL(1)
                count = count + 1
                MesLB.Items.Add(mes)
            Loop
        End If
        drL.Close()
        connL.Close()
        If (count = 1) Then
            MesLB.Items.Add("無")
        End If
        '************************
        SqlCmd = "select denyf from dbo.[user] where id='" & delid & "'"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        If (dr.HasRows) Then
            dr.Read()
            If (dr(0) = 0) Then
                CommUtil.ShowMsg(Me, "若此刪除之人為離職員工,事後請至帳號設定此人為離職狀態,以避免部門內簽核會再出現")
            End If
        End If
        dr.Close()
        conn.Close()
    End Sub
    'kkkkk
    Sub DDLItemAdd()
        Dim idstr As String
        DDLDelSIgnUser.Items.Clear()
        DDLDelSIgnUser.Items.Add("請選擇欲刪除之簽核人")
        DDLReplaceSignUser.Items.Clear()
        DDLReplaceSignUser.Items.Add("請選擇替代之簽核人")
        DDLReplaceSignUser.Items.Add("無需替代")
        SqlCmd = "select id,name,position from dbo.[user] where email<>'' order by denyf desc,branch,grp"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        If (dr.HasRows) Then
            Do While (dr.Read())
                idstr = dr(0) & " " & dr(1) & " " & dr(2)
                DDLDelSIgnUser.Items.Add(idstr)
            Loop
        End If
        dr.Close()
        conn.Close()
        SqlCmd = "select id,name,position from dbo.[user] where email<>'' and denyf<>1 order by branch,grp"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        If (dr.HasRows) Then
            Do While (dr.Read())
                idstr = dr(0) & " " & dr(1) & " " & dr(2)
                DDLReplaceSignUser.Items.Add(idstr)
            Loop
        End If
        dr.Close()
        conn.Close()
    End Sub
    Sub RecordSignFlowHistotyForUserChange(docnum As Long, reason As String, comment As String)
        Dim flowseq As Integer
        Dim signdate, agnname, sysid, sysname As String
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader

        signdate = Format(Now(), "yyyy/MM/dd HH:mm:ss")
        SqlCmd = "Select IsNull(Max(flowseq),0) from [dbo].[@XSPHT] where docentry=" & docnum
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        drL.Read()
        flowseq = dr(0) + 1
        drL.Close()
        connL.Close()
        agnname = ""
        sysid = "jetpmg"
        sysname = "系統通知"

        SqlCmd = "insert into [dbo].[@XSPHT] (docentry,uid,uname,flowseq,signdate,status,comment,agnname) " &
        "values(" & docnum & ",'" & sysid & "','" & sysname & "'," & flowseq &
        ",'" & signdate & "','" & reason & "','" & comment & "','" & agnname & "')"
        CommUtil.SqlSapExecute("ins", SqlCmd, connL)
        connL.Close()
    End Sub
    Protected Sub BtnSignAnalysis_Click(sender As Object, e As EventArgs) Handles BtnSignAnalysis.Click

        SignOffDataAnalysis()
    End Sub
    Protected Sub BtnSignModify_Click(sender As Object, e As EventArgs) Handles BtnSignModify.Click
        Dim str(), delid, delname, replaceid, replacename, replacepos, replaceemail, nextid() As String
        Dim connL, connL1, connL2 As New SqlConnection
        Dim drL, drL1, drL2 As SqlDataReader
        Dim samedelandreplace, isreplace, replacehassigned As Boolean
        Dim docstr, docstr1, docstr2, docstr3, docstr4, docstr5, comment, nextsignoffstr As String
        Dim idcount As Integer
        Dim replaceprop, replaceseq As Integer
        replaceprop = 100 'to make sure different from normal prop value if delprop<>replaceprop below
        idcount = 0
        nextid(0) = "end"
        comment = ""
        docstr = ""
        docstr1 = ""
        docstr2 = ""
        docstr3 = ""
        docstr4 = ""
        docstr5 = ""
        delname = ""
        replacename = ""
        replacepos = ""
        replaceemail = ""
        nextsignoffstr = ""
        samedelandreplace = False
        isreplace = False
        replacehassigned = False
        If (DDLDelSIgnUser.SelectedIndex = 0) Then
            CommUtil.ShowMsg(Me, "需選擇欲刪除(異動)之簽核人")
            Exit Sub
        Else
            str = Split(DDLDelSIgnUser.SelectedValue, " ")
            delid = str(0)
            delname = str(1)
        End If
        If (DDLReplaceSignUser.SelectedIndex = 0) Then
            CommUtil.ShowMsg(Me, "需選擇欲替代之簽核人")
            Exit Sub
        Else
            If (DDLReplaceSignUser.SelectedIndex = 1) Then
                replaceid = ""
            Else
                str = Split(DDLReplaceSignUser.SelectedValue, " ")
                replaceid = str(0)
            End If
        End If
        MesLB.Items.Clear()
        MesLB.Items.Add("以下若有 *********XXXXXXX******* 表示需手動處理 , 請特別注意")
        If (replaceid <> "") Then '有替代簽核者
            '以下為處理內定簽核表
            MesLB.Items.Add("")
            MesLB.Items.Add("簽核單之內建簽核人修正")
            SqlCmd = "select T0.sfid,T0.seq,T0.prop,T1.sfname,T0.num from [@XSPMT] T0 inner join [@XSFTT] T1 on T0.sfid=T1.sfid where uid='" & delid & "'"
            drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            Dim havereplaceid As Boolean
            Dim replacenum As Long
            If (drsap.HasRows) Then
                Do While (drsap.Read())
                    havereplaceid = False
                    samedelandreplace = False
                    SqlCmd = "select count(*) from [@XSPMT] T0 where sfid=" & drsap(0) & " and uid='" & replaceid & "'"
                    drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                    drL.Read()
                    If (drL(0) <> 0) Then
                        havereplaceid = True
                    End If
                    drL.Close()
                    connL.Close()
                    If (havereplaceid = False) Then '簽核表內無replaceid , 直接替換
                        SqlCmd = "Update [dbo].[@XSPMT] set uid='" & replaceid & "' where num=" & drsap(4)
                        CommUtil.SqlSapExecute("upd", SqlCmd, connL)
                        connL.Close()
                        docstr = docstr & drsap(0) & " "
                    Else
                        replaceprop = 100
                        SqlCmd = "select count(*) from [@XSPMT] T0 where prop=" & drsap(2) & " and sfid=" & drsap(0) & " and uid='" & replaceid & "'"
                        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                        drL.Read()
                        If (drL(0) > 0) Then
                            samedelandreplace = True
                        End If
                        drL.Close()
                        connL.Close()
                        If (samedelandreplace = False) Then
                            '  delid      replaceid     action
                            '    0            1         upd directly 
                            '    0            2         upd and delete replaceid 2
                            '    1            0         upd directly
                            '    1            2         upd delid to replaceid then delete replaceid 2
                            '    2            0         delete delid directly
                            '    2            1         delete delid directly
                            SqlCmd = "select prop,num from [@XSPMT] T0 where sfid=" & drsap(0) & " and uid='" & replaceid & "'"
                            drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                            If (drL.HasRows) Then
                                drL.Read()
                                replaceprop = drL(0)
                                replacenum = drL(1)
                            End If
                            drL.Close()
                            connL.Close()
                            If ((drsap(2) = 0 And replaceprop = 1) Or (drsap(2) = 1 And replaceprop = 0)) Then 'upd directly 
                                SqlCmd = "Update [dbo].[@XSPMT] set uid='" & replaceid & "' where num=" & drsap(4)
                                CommUtil.SqlSapExecute("upd", SqlCmd, connL)
                                connL.Close()
                                docstr4 = docstr4 & drsap(0) & " "
                            ElseIf ((drsap(2) = 0 And replaceprop = 2) Or (drsap(2) = 1 And replaceprop = 2)) Then 'upd and delete replaceid 2
                                SqlCmd = "Update [dbo].[@XSPMT] set uid='" & replaceid & "' where num=" & drsap(4)
                                CommUtil.SqlSapExecute("upd", SqlCmd, connL)
                                connL.Close()
                                'docstr5 = docstr5 & drsap(0) & " "
                                SqlCmd = "delete from [dbo].[@XSPMT] where num=" & replacenum
                                CommUtil.SqlSapExecute("del", SqlCmd, connL)
                                connL.Close()
                                docstr2 = docstr2 & drsap(0) & " "
                            ElseIf ((drsap(2) = 2 And replaceprop = 0) Or (drsap(2) = 2 And replaceprop = 1)) Then 'delete delid directly
                                SqlCmd = "delete from [dbo].[@XSPMT] where num=" & drsap(4)
                                CommUtil.SqlSapExecute("del", SqlCmd, connL)
                                connL.Close()
                                docstr1 = docstr1 & drsap(0) & " "
                            Else
                                CommUtil.ShowMsg(Me, "發生條件例外情況")
                            End If
                        Else
                            '  delid      replaceid     action
                            '     0           0         del delid
                            '     1           1         del delid 應該是無此狀況
                            '     2           2         del delid
                            'MesLB.Items.Add("替代id :" & replaceid & " 已存在於單據:" & drsap(0) & "(" & drsap(3) & "), 已直接刪除原id:" & delid)
                            SqlCmd = "delete from [dbo].[@XSPMT] where num=" & drsap(4)
                            CommUtil.SqlSapExecute("del", SqlCmd, connL)
                            connL.Close()
                            docstr3 = docstr3 & drsap(0) & " "
                            If (drsap(2) = 0) Then
                                SqlCmd = "Update [dbo].[@XSPMT] set seq = seq - 1 where sfid=" & drsap(0) & " and seq > " & drsap(1) & " and prop = 0"
                                CommUtil.SqlSapExecute("upd", SqlCmd, connL)
                                connL.Close()
                            End If
                        End If
                    End If
                Loop
            End If
            drsap.Close()
            connsap.Close()
            If (docstr <> "" Or docstr1 <> "" Or docstr2 <> "" Or docstr3 <> "" Or docstr4 <> "") Then
                If (docstr <> "") Then
                    MesLB.Items.Add("===>單據種類: " & docstr & "內本無替代者, 現已用替代者替代完畢")
                End If
                If (docstr1 <> "") Then
                    MesLB.Items.Add("===>單據種類: " & docstr1 & "因原簽核人(知悉)與替代人同存在於內定簽核表中,故只刪除原簽核人")
                End If
                If (docstr2 <> "") Then
                    MesLB.Items.Add("===>單據種類: " & docstr2 & "因原簽核人與替代人(知悉)同存在於內定簽核表中,故刪除原替代人, 再用替代人代替原簽核人")
                End If
                If (docstr3 <> "") Then
                    MesLB.Items.Add("===>單據種類: " & docstr3 & "因原簽核人與替代人(知悉)同存在於內定簽核表中,且同簽核屬性,故刪除原替代人")
                End If
                If (docstr4 <> "") Then
                    MesLB.Items.Add("===>單據種類: " & docstr4 & "因原簽核人與替代人(知悉)同存在於內定簽核表中,現已用替代者替代完畢, 並保留原替代者簽核設定")
                End If
            Else
                MesLB.Items.Add("===>無需替代")
            End If

            MesLB.Items.Add("")
            MesLB.Items.Add("自訂簽核單之自訂簽核人組別修正")


            '以下為處理在線簽核表單
            docstr = ""
            docstr1 = ""
            docstr2 = ""
            docstr3 = ""
            docstr4 = ""
            MesLB.Items.Add("")
            MesLB.Items.Add("在線表單簽核人修正")
            SqlCmd = "select name,position,email,seq from dbo.[user] where uid ='" & replaceid & "'"
            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
            If (dr.HasRows) Then
                Do While (dr.Read())
                    replacename = dr(0)
                    replacepos = dr(1)
                    replaceemail = dr(2)
                Loop
            End If
            dr.Close()
            conn.Close()

            ' 下述seq=1(發起者) 不作替換動作 (若有,在List Box會提示處理(因若有是被recall或退回 , 需以作廢處理)==> 還是替換 2025/1/14
            '若簽核記錄中本已有 replaceid 存在 , 則也不替換

            'SqlCmd = "select docentry,num,seq,status from [@XSPWT] T0 where uid='" & delid & "' and (status = 0 or status = 1) and seq <> 1"

            SqlCmd = "select docentry,num,seq,status,signprop from [@XSPWT] T0 where uid='" & delid & "' and (status = 0 or status = 1)"
            drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (drsap.HasRows) Then
                Do While (drsap.Read())
                    havereplaceid = False
                    samedelandreplace = False
                    replacehassigned = False
                    SqlCmd = "select signprop,num,seq,status from [@XSPWT] T0 where docentry=" & drsap(0) & " and uid='" & replaceid & "'"
                    drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                    If (drL.HasRows) Then
                        drL.Read()
                        replaceprop = drL(0)
                        replacenum = drL(1)
                        replaceseq = drL(2)
                        havereplaceid = True
                        If (drsap(3) <> 0 And drsap(3) <> 1 And drsap(3) <> 105) Then
                            replacehassigned = True
                        End If
                    End If
                    drL.Close()
                    connL.Close()

                    If (havereplaceid = False) Then '簽核表內無replaceid , 直接替換
                        SqlCmd = "Update [dbo].[@XSPWT] set uid='" & replaceid & "',uname='" & replacename & "',upos='" & replacepos & "'," &
                                "emailadd='" & replaceemail & "' where num = " & drsap(1)
                        CommUtil.SqlSapExecute("upd", SqlCmd, connL)
                        connL.Close()
                        If (drsap(3) = 1) Then
                            isreplace = True
                        End If
                        docstr = docstr & drsap(0) & " "
                        comment = "原簽核人:" & delname & "換成: " & replacename
                        RecordSignFlowHistotyForUserChange(drsap(0), "人員異動", comment)
                    Else
                        replaceprop = 100
                        SqlCmd = "select count(*) from [@XSPWT] T0 where signprop=" & drsap(4) & " and docentry=" & drsap(0) & " and uid='" & replaceid & "'"
                        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                        drL.Read()
                        If (drL(0) > 0) Then
                            samedelandreplace = True
                        End If
                        drL.Close()
                        connL.Close()
                        'If (drL1(0) = 0) Then '原簽核無replaceid
                        If (samedelandreplace = False) Then
                            '  delid      replaceid             action
                            '    0            1                 upd directly 
                            '    0            2                 upd and delete replaceid 2
                            '    1            0                 upd directly
                            '    1            2                 upd delid to replaceid then delete replaceid 2
                            '    2            0                 delete delid directly
                            '    2            1                 delete delid directly
                            If ((drsap(4) = 0 And replaceprop = 1) Or (drsap(4) = 1 And replaceprop = 0)) Then 'upd directly 
                                SqlCmd = "Update [dbo].[@XSPWT] set uid='" & replaceid & "',uname='" & replacename & "',upos='" & replacepos & "'," &
                                "emailadd='" & replaceemail & "' where num = " & drsap(1)
                                CommUtil.SqlSapExecute("upd", SqlCmd, connL)
                                connL.Close()
                                docstr4 = docstr4 & drsap(0) & " "
                                comment = "原簽核人:" & delname & "換成: " & replacename & " 並保留原替代人設定"
                                RecordSignFlowHistotyForUserChange(drsap(0), "人員異動", comment)
                                If (drsap(3) = 1) Then
                                    isreplace = True
                                End If
                            ElseIf ((drsap(4) = 0 And replaceprop = 2) Or (drsap(4) = 1 And replaceprop = 2)) Then 'upd and delete replaceid 2
                                SqlCmd = "Update [dbo].[@XSPWT] set uid='" & replaceid & "',uname='" & replacename & "',upos='" & replacepos & "'," &
                                "emailadd='" & replaceemail & "' where num = " & drsap(1)
                                CommUtil.SqlSapExecute("upd", SqlCmd, connL)
                                connL.Close()
                                'docstr5 = docstr5 & drsap(0) & " "
                                SqlCmd = "delete from [dbo].[@XSPWT] where num=" & replacenum
                                CommUtil.SqlSapExecute("del", SqlCmd, connL)
                                connL.Close()
                                docstr2 = docstr2 & drsap(0) & " "
                                comment = "原簽核人:" & delname & "換成: " & replacename & " 並刪除原替代人(知悉)"
                                RecordSignFlowHistotyForUserChange(drsap(0), "人員異動", comment)
                                If (drsap(3) = 1) Then
                                    isreplace = True
                                End If
                            ElseIf ((drsap(4) = 2 And replaceprop = 0) Or (drsap(4) = 2 And replaceprop = 1)) Then 'delete delid directly
                                SqlCmd = "delete from [dbo].[@XSPWT] where num=" & drsap(1)
                                CommUtil.SqlSapExecute("del", SqlCmd, connL)
                                connL.Close()
                                docstr1 = docstr1 & drsap(0) & " "
                                comment = "替代簽核人:" & replacename & "原已存在 , 刪除原簽核人(知悉)"
                                RecordSignFlowHistotyForUserChange(drsap(0), "人員異動", comment)
                            Else
                                CommUtil.ShowMsg(Me, "發生條件例外情況")
                            End If
                        Else
                            '  delid      replaceid     action
                            '     0           0         del delid
                            '     1           1         del delid 應該是無此狀況
                            '     2           2         del delid
                            ' delete 替代者 , upd 刪除者為替代者 , 可包括 seq=1 狀況 (若delete 刪除者 , 保留替代者 , 則需再考慮seq=1 及 <>1 情況)
                            SqlCmd = "delete from [dbo].[@XSPWT] where num=" & replacenum
                            CommUtil.SqlSapExecute("del", SqlCmd, connL)
                            connL.Close()
                            SqlCmd = "Update [dbo].[@XSPWT] set uid='" & replaceid & "',uname='" & replacename & "',upos='" & replacepos & "'," &
                                     "emailadd='" & replaceemail & "' where num = " & drsap(1)
                            CommUtil.SqlSapExecute("upd", SqlCmd, connL)
                            connL.Close()
                            If (drsap(4) = 0) Then
                                SqlCmd = "Update [dbo].[@XSPWT] set seq = seq - 1 where docentry=" & drsap(0) & " and seq > " & replaceseq
                                CommUtil.SqlSapExecute("upd", SqlCmd, connL2)
                                connL2.Close()
                            End If

                            If (drsap(3) = 1) Then
                                SqlCmd = "Update [dbo].[@XSPWT] set status = 1 where docentry=" & drsap(0) & " and seq = " & drsap(2)
                                CommUtil.SqlSapExecute("upd", SqlCmd, connL2)
                                connL2.Close()
                                '要通知修改完後關卡是誰
                                SqlCmd = "select uid from [@XSPWT] T0 where docentry=" & drsap(3) & " and seq = " & drsap(2)
                                drL2 = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL2)
                                If (drL2.HasRows) Then
                                    drL2.Read()
                                    If (InStr(nextsignoffstr, drL2(0)) = 0) Then
                                        nextsignoffstr = nextsignoffstr & drL2(0) & " "
                                        nextid(idcount) = drL2(0)
                                        nextid(idcount + 1) = "end"
                                        idcount = idcount + 1
                                        'CommSignOff.ReplaceSignOffInform(Application("http"), drL2(0), "原簽核人被刪除重發簽核")
                                    End If
                                End If
                                drL2.Close()
                                connL2.Close()
                                comment = "原簽核人:" & delname & "換成: " & replacename & " 並刪除原替代人設定(同屬性)"
                                RecordSignFlowHistotyForUserChange(drsap(0), "人員異動", comment)
                            End If
                            docstr3 = docstr3 & drsap(0) & " "
                            If (drsap(3) = 1) Then
                                isreplace = True
                            End If
                        End If
                    End If
                Loop
            End If
            drsap.Close()
            connsap.Close()
            If (docstr <> "" Or docstr1 <> "" Or docstr2 <> "" Or docstr3 <> "" Or docstr4 <> "") Then
                If (docstr <> "") Then
                    MesLB.Items.Add("===>單據種類: " & docstr & "內本無替代者, 現已用替代者替代完畢")
                End If
                If (docstr1 <> "") Then
                    MesLB.Items.Add("===>單據種類: " & docstr1 & "因原簽核人(知悉)與替代人同存在於內定簽核表中,故只刪除原簽核人")
                End If
                If (docstr2 <> "") Then
                    MesLB.Items.Add("===>單據種類: " & docstr2 & "因原簽核人與替代人(知悉)同存在於內定簽核表中,故刪除原替代人, 再用替代人代替原簽核人")
                End If
                If (docstr3 <> "") Then
                    MesLB.Items.Add("===>單據種類: " & docstr3 & "因原簽核人與替代人(知悉)同存在於內定簽核表中,且同簽核屬性,故替代刪除人並刪除原替代人簽核")
                End If
                If (docstr4 <> "") Then
                    MesLB.Items.Add("===>單據種類: " & docstr4 & "因原簽核人與替代人(知悉)同存在於內定簽核表中,現已用替代者替代完畢, 並保留原替代者簽核設定")
                End If
            Else
                MesLB.Items.Add("===>無需替代")
            End If

            MesLB.Items.Add("")
            MesLB.Items.Add("未送審單據建立人替代:")
            docstr = ""
            SqlCmd = "SELECT T1.docnum " &
                 "FROM dbo.[@XASCH] T1 Inner Join [@XSFTT] T2 On T1.sfid=T2.sfid " &
                 "where T1.sid='" & delid & "' and (T1.status='E' or T1.status='D') " &
                 " order by T1.sfid,T1.docnum desc"
            drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
            If (drL.HasRows) Then
                Do While (drL.Read())
                    SqlCmd = "Update [dbo].[@XASCH] set sid = '" & replaceid & "',sname='" & replacename & "' " &
                            "where docentry=" & drL(0)
                    CommUtil.SqlSapExecute("upd", SqlCmd, connL2)
                    connL2.Close()
                    docstr = docstr & drL(0) & " "
                    isreplace = True
                    comment = "原簽核建立人:" & delname & "換成: " & replacename
                    RecordSignFlowHistotyForUserChange(drL(0), "人員異動", comment)
                Loop
            End If
            drL.Close()
            connL.Close()
            If (docstr <> "") Then
                MesLB.Items.Add("單據編號: " & docstr & "===>未送審單據建立人替代完畢")
            Else
                MesLB.Items.Add("===>無需替代")
            End If
            If (isreplace = True) Then
                CommSignOff.ReplaceSignOffInform(Application("http"), replaceid, "替換簽核")
            End If
            idcount = 0
            Do While (nextid(idcount) <> "end")
                CommSignOff.ReplaceSignOffInform(Application("http"), nextid(idcount), "原簽核人被刪除重發簽核")
                idcount = idcount + 1
            Loop
            'docstr = ""
            'SqlCmd = "select docentry from [@XSPWT] T0 where uid='" & delid & "' and (status = 0 or status = 1) and seq = 1"
            'drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
            'If (drsap.HasRows) Then
            '    drL.Read()
            '    docstr = docstr & drL(0) & " "
            'End If
            'drL.Close()
            'connL.Close()

            'If (docstr <> "") Then
            '    'MesLB.ForeColor = System.Drawing.Color.Red
            '    MesLB.Items.Add("===>單據編號:  " & docstr & "發現送審者需替代,請手動處理")
            'End If
        Else
            ''若有seq=1(發起者)及 prop=1(歸檔者)處理者 ,會提示要先處理 , 才能做刪除指定簽核user
            'SqlCmd = "select seq,prop from [@XSPWT] T0 where uid='" & delid & "' and (status = 0 or status = 1) and (seq = 1 or prop = 1)"
            'drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            'If (drsap.HasRows) Then

            '    Exit Sub
            'End If
            'drsap.Close()
            'connsap.Close()

            'SqlCmd = "delete from [dbo].[@XSPMT] where prop <> 0 and uid='" & delid & "'"
            'CommUtil.SqlSapExecute("del", SqlCmd, connsap)
            'connsap.Close()

            SqlCmd = "select sfid,seq,prop from [@XSPMT] T0 where uid='" & delid & "'"
            drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (drsap.HasRows) Then
                Do While (drsap.Read())
                    If (drsap(2) <> 1) Then
                        SqlCmd = "delete from [dbo].[@XSPMT] where uid='" & delid & "' and sfid=" & drsap(0) & " and prop=" & drsap(2)
                        CommUtil.SqlSapExecute("del", SqlCmd, connL)
                        connL.Close()
                        If (drsap(2) = 0) Then
                            SqlCmd = "Update [dbo].[@XSPMT] set seq = seq - 1 where sfid=" & drsap(0) & " and seq > " & drsap(1) & " and prop = 0"
                            CommUtil.SqlSapExecute("upd", SqlCmd, connL)
                            connL.Close()
                        End If
                        docstr = docstr & drsap(0) & " "
                    Else
                        docstr1 = docstr1 & drsap(0) & " "
                    End If
                Loop
            End If
            drsap.Close()
            connsap.Close()
            MesLB.Items.Add("內建簽核人修正:")
            If (docstr <> "" Or docstr1 <> "") Then
                If (docstr <> "") Then
                    MesLB.Items.Add("===>單據種類: " & docstr & "選擇的刪除人刪除完畢")
                End If
                If (docstr1 <> "") Then
                    MesLB.Items.Add("******===>單據種類: " & docstr1 & "有此刪除人當歸檔者,請手動處理(因為單據最後處理者,無替代者,直接刪除--不妥********")
                End If
            Else
                MesLB.Items.Add("===>無需刪除修正")
            End If
            '以下為處理在線簽核表單
            'seq=1(發起者)及 prop=1(歸檔者)不處理,若有,在List Box , 會提示處理
            docstr = ""
            nextsignoffstr = ""
            idcount = 0
            SqlCmd = "select seq,prop,num,docentry,status from [@XSPWT] T0 where uid='" & delid & "' and (status = 0 or status = 1) and seq <> 1 and prop <> 1"
            drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (drsap.HasRows) Then
                Do While (drsap.Read())
                    SqlCmd = "Update [dbo].[@XSPWT] set seq = seq - 1 where docentry=" & drsap(3) & " and seq > " & drsap(0)
                    CommUtil.SqlSapExecute("upd", SqlCmd, connL)
                    connL.Close()
                    SqlCmd = "delete from [dbo].[@XSPWT] where num=" & drsap(2)
                    CommUtil.SqlSapExecute("del", SqlCmd, connL)
                    connL.Close()
                    docstr = docstr & drsap(3) & " "
                    comment = "原簽核人:" & delname & "刪除 , 且無替代人"
                    RecordSignFlowHistotyForUserChange(drsap(3), "人員異動", comment)
                    If (drsap(4) = 1) Then
                        SqlCmd = "Update [dbo].[@XSPWT] set status = 1 where docentry=" & drsap(3) & " and seq = " & drsap(0)
                        CommUtil.SqlSapExecute("upd", SqlCmd, connL)
                        connL.Close()
                        '要通知修改完後關卡是誰
                        SqlCmd = "select uid from [@XSPWT] T0 where docentry=" & drsap(3) & " and seq = " & drsap(0)
                        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                        If (drL.HasRows) Then
                            drL.Read()
                            'CommSignOff.ReplaceSignOffInform(Application("http"), drL(0), "原簽核人被刪除重啟簽核")
                            If (InStr(nextsignoffstr, drL(0)) = 0) Then
                                nextsignoffstr = nextsignoffstr & drL(0) & " "
                                nextid(idcount) = drL(0)
                                nextid(idcount + 1) = "end"
                                idcount = idcount + 1
                                'CommSignOff.ReplaceSignOffInform(Application("http"), drL2(0), "原簽核人被刪除重發簽核")
                            End If
                        End If
                        drL.Close()
                        connL.Close()
                    End If
                Loop
            End If
            drsap.Close()
            connsap.Close()
            MesLB.Items.Add("")
            MesLB.Items.Add("在線表單簽核人修正")
            If (docstr <> "") Then
                MesLB.Items.Add("單據編號: " & docstr & "===>因無替代人 , 故選擇之刪除人刪除完畢")
            Else
                MesLB.Items.Add("===>無單據需刪除")
            End If
            SqlCmd = "select seq,prop,num,docentry,status from [@XSPWT] T0 where uid='" & delid & "' and (status = 0 or status = 1) and (seq = 1 or prop = 1)"
            drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (drsap.HasRows) Then
                MesLB.Items.Add("")
                MesLB.Items.Add("***********下列為刪除者為再送審或歸檔者(需手動處理) :***********")
                Do While (drsap.Read())
                    If (drsap(0) = 1) Then
                        MesLB.Items.Add("單號:" & drsap(3) & "===>" & delname & " 為再送審者")
                    ElseIf (drsap(1) = 1) Then
                        MesLB.Items.Add("單號:" & drsap(3) & "===>" & delname & " 為歸檔者")
                    End If
                Loop
            End If
            drsap.Close()
            connsap.Close()
            MesLB.Items.Add("")
            MesLB.Items.Add("***********未送審單據需手動處理List:***********")
            docstr = ""
            SqlCmd = "SELECT T1.docnum " &
                 "FROM dbo.[@XASCH] T1 Inner Join [@XSFTT] T2 On T1.sfid=T2.sfid " &
                 "where T1.sid='" & delid & "' and (T1.status='E' or T1.status='D') " &
                 " order by T1.sfid,T1.docnum desc"
            drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
            If (drL.HasRows) Then
                Do While (drL.Read())
                    docstr = docstr & drL(0) & " "
                Loop
            End If
            drL.Close()
            connL.Close()
            If (docstr <> "") Then
                MesLB.Items.Add("單據編號: " & docstr & "===>***********未送審單據因無替代人 ,需手動處理***********")
            Else
                MesLB.Items.Add("===>無未送審單據需處理")
            End If
            idcount = 0
            Do While (nextid(idcount) <> "end")
                CommSignOff.ReplaceSignOffInform(Application("http"), nextid(idcount), "原簽核人被刪除重發簽核")
                idcount = idcount + 1
            Loop
        End If
        MesLB.Items.Add("")
        MesLB.Items.Add("以上為處理結果, 請特別注意需手動處理部份")
    End Sub
End Class