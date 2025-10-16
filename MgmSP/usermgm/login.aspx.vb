Imports System.Data
Imports System.Data.SqlClient

Partial Public Class login
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public SqlCmd As String
    Public dr As SqlDataReader
    Public conn, connsap As New SqlConnection

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Dim count As Integer
        If (Not IsPostBack) Then
            'InitSAPSQLConnection(ServerText.Text, "")
            'DDLServer.Visible = True
            'DDLWhs.Visible = True
            'DDLServer.Items.Clear()
            'DDLServer.Items.Add("請選擇SAP資料庫")

            'SqlCmd = "SELECT name, database_id, create_date FROM sys.databases"
            'myCommand = New SqlCommand(SqlCmd, connsap)
            'dr = myCommand.ExecuteReader()
            'Do While (dr.Read())
            '    If (dr(0) <> "master" And dr(0) <> "tempdb" And dr(0) <> "model" And dr(0) <> "msdb" And dr(0) <> "SBO-COMMON") Then
            '        DDLServer.Items.Add(dr(0))
            '    End If
            'Loop
            'dr.Close()
            'CloseSAPSQLConnection()
            'DDLWhs.Items.Clear()


            'count = 0
            'SqlCmd = "SELECT T0.[WhsCode], T0.[WhsName] FROM OWHS T0 order by T0.WhsCode"
            'dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)

            DDLWhs.Items.Clear()
            DDLWhs.Items.Add("C01 ICT")
            DDLWhs.Items.Add("C02 AOI")
            DDLWhs.SelectedIndex = 1
            'DDLWhs.Items.Add("請選擇倉別")
            'Do While (dr.Read())
            'DDLWhs.Items.Add(dr(0) & " " & dr(1))
            'count = count + 1
            'Loop
            ''''''''
            'If (count >= 4) Then
            'DDLWhs.SelectedIndex = 4
            'Else
            'DDLWhs.SelectedIndex = 2
            'End If
            ''''''''
            'dr.Close()
            'connsap.Close()
            Dim actmode, uid, agnid, inchargeid As String
            Dim docnum As Long
            Dim formstatusindex, formtypeindex, sfid As Integer
            Dim str(), docstatus As String
            actmode = Request.QueryString("actmode")
            Session("actmode") = ""
            If (actmode = "signoff" Or actmode = "single_signoff" Or actmode = "todoitem" Or actmode = "informtraceperson") Then
                uid = Request.QueryString("uid")
                docnum = Request.QueryString("docnum")
                formstatusindex = Request.QueryString("formsatusindex")
                formtypeindex = Request.QueryString("formtypeindex")
                docstatus = Request.QueryString("status")
                sfid = Request.QueryString("sfid")
                agnid = Request.QueryString("agnid")
                inchargeid = Request.QueryString("inchargeid")
                SqlCmd = "Select id,name,pwd,email,ttl,area,denyf,sapid,sappwd,grp,branch From dbo.[User] where dbo.[User].id='" & uid & "'"
                dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                If (dr.HasRows) Then
                    dr.Read()
                    Session("s_id") = uid
                    Session("s_name") = dr("name")
                    Session("sapid") = dr("sapid")
                    Session("sappwd") = dr("sappwd")
                    Session("branch") = dr("branch")
                    Session("grp") = dr("grp")
                End If
                dr.Close()
                conn.Close()
                Session("usingserver") = "192.168.1.31"
                'Session("usingdb") = "JTSTD"
                Session("usingdb") = "JTTST1"
                Session("actmode") = actmode
                'If (actmode = "todoitem") Then
                '    Session("actmode") = "todoitem"
                'Else
                '    Session("actmode") = "signoff" '表示從email進入
                'End If
                str = Split(DDLWhs.SelectedValue, " ")
                Session("usingwhsfull") = DDLWhs.SelectedValue
                Session("usingwhs") = str(0)
                If (actmode = "signoff" Or actmode = "single_signoff") Then
                    Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&actmode=" & actmode & "&uid=" & uid &
                                  "&status=" & docstatus & "&formtypeindex=" & formtypeindex & "&formstatusindex=" & formstatusindex &
                                  "&docnum=" & docnum & "&sfid=" & sfid & "&agnid=" & agnid)
                ElseIf (actmode = "todoitem") Then
                    Response.Redirect("~/signoff/signofftodo.aspx?smid=sg&smode=6&actstr=todoitem&uid=" & uid & "&inchargeid=" & inchargeid)
                ElseIf (actmode = "informtraceperson") Then
                    Response.Redirect("~/signoff/signofftodo.aspx?smid=sg&smode=6&actstr=informtraceperson&uid=" & uid & "&inchargeid=" & inchargeid &
                                      "&num=" & Request.QueryString("num"))
                End If
            End If
        End If
    End Sub

    Protected Sub loginbtn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles loginbtn.Click
        Dim str() As String
        If (idtxt.Text = "") Then
            errmsg.Text = "沒輸入帳號"
            Exit Sub
        End If
        If (pwdtxt.Text = "") Then
            errmsg.Text = "沒輸入密碼"
            Exit Sub
        End If
        If (DDLServer.SelectedIndex = 0) Then
            errmsg.Text = "沒選擇資料庫"
            Exit Sub
        End If
        SqlCmd = "Select id,name,pwd,email,ttl,area,denyf,sapid,sappwd,grp,branch From dbo.[User] where dbo.[User].id='" & idtxt.Text & "'"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        If (dr.HasRows) Then
            dr.Read()
            If (dr("pwd") <> pwdtxt.Text) Then
                errmsg.Text = "密碼錯誤"
            Else
                If (dr("denyf") = 1) Then
                    errmsg.Text = "停用帳號"
                Else
                    Session("s_id") = dr("id")
                    Session("s_name") = dr("name")
                    'Session("usingdb") = "JTSTD"
                    Session("usingdb") = "JTTST1"
                    str = Split(DDLWhs.SelectedValue, " ")
                    Session("usingwhsfull") = DDLWhs.SelectedValue
                    Session("usingwhs") = str(0)
                    Session("sapid") = dr("sapid")
                    Session("sappwd") = dr("sappwd")
                    Session("grp") = dr("grp")
                    Session("branch") = dr("branch")
                    Session("usingserver") = "192.168.1.31" 'ServerText.Text
                    dr.Close()

                    Response.Redirect("~/index.aspx?smid=index")
                End If
            End If
        Else
            errmsg.Text = "無此帳號"
        End If
        conn.Close()
    End Sub
    'object delete
    'Protected Sub ServerText_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ServerText.TextChanged
    '    DDLServer.Items.Clear()
    '    InitSAPSQLConnection(ServerText.Text, "")
    '    DDLServer.Items.Add("請選擇SAP資料庫")
    '    SqlCmd = "SELECT name, database_id, create_date FROM sys.databases"
    '    myCommand = New SqlCommand(SqlCmd, connsap)
    '    dr = myCommand.ExecuteReader()
    '    Do While (dr.Read())
    '        If (dr(0) <> "master" And dr(0) <> "tempdb" And dr(0) <> "model" And dr(0) <> "msdb" And dr(0) <> "SBO-COMMON") Then
    '            DDLServer.Items.Add(dr(0))
    '        End If
    '    Loop
    '    '''''''''
    '    DDLServer.SelectedIndex = 1
    '    '''''''''
    '    dr.Close()
    '    CloseSAPSQLConnection()
    'End Sub

    'object delete
    'Protected Sub DDLServer_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles DDLServer.SelectedIndexChanged
    '    If (DDLServer.SelectedIndex <> 0) Then
    '        InitSAPSQLConnection(ServerText.Text, DDLServer.SelectedValue)
    '        SqlCmd = "SELECT T0.[WhsCode], T0.[WhsName] FROM OWHS T0 order by T0.WhsCode"
    '        myCommand = New SqlCommand(SqlCmd, connsap)
    '        dr = myCommand.ExecuteReader()
    '        DDLWhs.Items.Clear()
    '        DDLWhs.Items.Add("請選擇倉別")
    '        Do While (dr.Read())
    '            DDLWhs.Items.Add(dr(0) & " " & dr(1))
    '        Loop
    '        ''''''''
    '        DDLWhs.SelectedIndex = 2
    '        ''''''''
    '        dr.Close()
    '        CloseSAPSQLConnection()
    '    Else
    '        DDLWhs.Items.Clear()
    '    End If
    'End Sub
End Class