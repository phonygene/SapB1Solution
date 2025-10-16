Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Net.Mail
Imports itextsharp.text
Imports iTextSharp.text.pdf
Imports AjaxControlToolkit
Public Class CommUtil
    Inherits System.Web.UI.Page
    Public oCompany As New SAPbobsCOM.Company
    Public connsap, conn As New SqlConnection
    Public SqlCmd As String
    Public dr, dr1, drsap As SqlDataReader

    Public Sub ConfirmBox(Sender As Object, ByVal Message As String)
        Dim sScript As String
        Dim sMessage As String
        sMessage = Strings.Replace(Message, "'", "\'") '處理單引號
        sMessage = Strings.Replace(sMessage, vbNewLine, "\n") '處理換行
        sScript = String.Format("confirm('{0}');", sMessage)
        ScriptManager.RegisterStartupScript(Sender, Sender.GetType(), "alert", sScript, True)
    End Sub
    Public Sub ShowMsg(Sender As Object, ByVal Message As String)
        Dim sScript As String
        Dim sMessage As String
        sMessage = Strings.Replace(Message, "'", "\'") '處理單引號
        sMessage = Strings.Replace(sMessage, vbNewLine, "\n") '處理換行
        sScript = String.Format("alert('{0}');", sMessage)
        ScriptManager.RegisterStartupScript(Sender, Sender.GetType(), "alert", sScript, True)
    End Sub
    Public Sub InitLocalSQLConnection(conn As SqlConnection)
        conn.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString
        conn.Open()
    End Sub

    Public Sub InitSAPSQLConnection(conn As SqlConnection)
        conn.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("SapSQLConnection").ConnectionString
        conn.Open()
    End Sub

    Public Function InitSAPConnection(ByVal DestIP As String, ByVal HostName As String, sapid As String, sappwd As String) As Long
        oCompany.Server = DestIP
        oCompany.CompanyDB = HostName
        oCompany.UserName = sapid
        oCompany.Password = sappwd
        oCompany.UseTrusted = False
        oCompany.DbUserName = "sa"
        oCompany.DbPassword = "sap19690123"
        oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
        oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005
        InitSAPConnection = oCompany.Connect
    End Function

    Public Sub CloseSAPConnection()
        oCompany.Disconnect()
    End Sub

    Public Function SelectLocalSqlUsingDr(SqlCmd As String, conn As SqlConnection)
        Dim dr As SqlDataReader
        conn.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString
        conn.Open()
        Dim myCommand As SqlCommand
        myCommand = New SqlCommand(SqlCmd, conn)
        dr = myCommand.ExecuteReader()
        Return dr
    End Function

    Public Function SelectSapSqlUsingDr(SqlCmd As String, conn As SqlConnection)
        Dim dr As SqlDataReader
        conn.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("SapSQLConnection").ConnectionString
        conn.Open()
        Dim myCommand As SqlCommand
        myCommand = New SqlCommand(SqlCmd, conn)
        dr = myCommand.ExecuteReader()
        Return dr
    End Function

    Public Function SelectSqlUsingDr(dr As SqlDataReader, SqlCmd As String, conn As SqlConnection)
        Dim myCommand As SqlCommand
        myCommand = New SqlCommand(SqlCmd, conn)
        dr = myCommand.ExecuteReader()
        Return dr
    End Function

    Public Function SelectLocalSqlUsingDataSet(ds As DataSet, SqlCmd As String, conn As SqlConnection)
        Dim myCommand As SqlCommand
        conn.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString
        conn.Open()
        myCommand = New SqlCommand(SqlCmd, conn)
        Dim da1 As New SqlDataAdapter(SqlCmd, conn)
        da1.Fill(ds)
        Return ds
    End Function

    Public Function SelectSapSqlUsingDataSet(ds As DataSet, SqlCmd As String, conn As SqlConnection)
        Dim myCommand As SqlCommand
        myCommand = New SqlCommand(SqlCmd, conn)
        conn.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("SapSQLConnection").ConnectionString
        conn.Open()
        Dim da1 As New SqlDataAdapter(SqlCmd, conn)
        da1.Fill(ds)
        Return ds
    End Function

    Public Function SelectSqlUsingDataSet(ds As DataSet, SqlCmd As String, conn As SqlConnection)
        Dim myCommand As SqlCommand
        myCommand = New SqlCommand(SqlCmd, conn)
        Dim da1 As New SqlDataAdapter(SqlCmd, conn)
        da1.Fill(ds)
        Return ds
    End Function

    Public Function SqlExecute(sqltype As String, SqlCmd As String, conn As SqlConnection)
        Dim myCommand As SqlCommand
        Dim count As Integer
        myCommand = New SqlCommand(SqlCmd, conn)
        count = myCommand.ExecuteNonQuery()
        If (count = 0) Then
            If (sqltype = "upd") Then
                ShowMsg(Me, "更新狀態失敗")
            ElseIf (sqltype = "del") Then
                ShowMsg(Me, "刪除失敗")
            ElseIf (sqltype = "ins") Then
                ShowMsg(Me, "新增失敗")
            Else
                ShowMsg(Me, "執行之sqltype可能寫錯")
            End If
            SqlExecute = False
        Else
            SqlExecute = True
        End If
    End Function

    Public Function SqlLocalExecute(sqltype As String, SqlCmd As String, conn As SqlConnection)
        Dim myCommand As SqlCommand
        Dim count As Integer
        conn.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString
        conn.Open()
        myCommand = New SqlCommand(SqlCmd, conn)
        count = myCommand.ExecuteNonQuery()
        If (count = 0) Then
            If (sqltype = "upd") Then
                ShowMsg(Me, "更新狀態失敗")
            ElseIf (sqltype = "del") Then
                ShowMsg(Me, "刪除失敗")
            ElseIf (sqltype = "ins") Then
                ShowMsg(Me, "新增失敗")
            ElseIf (sqltype = "alter") Then
                ShowMsg(Me, "alter table reset 失敗")
            Else
                ShowMsg(Me, "執行之sqltype可能寫錯")
            End If
            SqlLocalExecute = False
        Else
            SqlLocalExecute = True
        End If
    End Function

    Public Function SqlSapExecute(sqltype As String, SqlCmd As String, conn As SqlConnection)
        Dim myCommand As SqlCommand
        Dim count As Integer
        conn.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("SapSQLConnection").ConnectionString
        conn.Open()
        myCommand = New SqlCommand(SqlCmd, conn)
        count = myCommand.ExecuteNonQuery()
        If (count = 0) Then
            If (sqltype = "upd") Then
                ShowMsg(Me, "更新狀態失敗")
            ElseIf (sqltype = "del") Then
                ShowMsg(Me, "刪除失敗")
            ElseIf (sqltype = "ins") Then
                ShowMsg(Me, "新增失敗")
            Else
                ShowMsg(Me, "執行之sqltype可能寫錯")
            End If
            SqlSapExecute = False
        Else
            SqlSapExecute = True
        End If
    End Function
    Function DisableObjectByPermission(sender As Object, perms As String, kp As String)
        If (InStr(perms, kp)) Then '如kp 是perms一部分的話
            sender.Enabled = True
            DisableObjectByPermission = True
        Else
            sender.Enabled = False
            DisableObjectByPermission = False
        End If
    End Function

    Function GetAssignRight(ByVal sysid As String, s_id As String)
        Dim perms As String
        Dim SqlCmd As String
        Dim conn As New SqlConnection
        Dim dr As SqlDataReader
        perms = ""
        SqlCmd = "Select T0.permission From dbo.[user_permissionnew] T0 where T0.id='" & s_id & "' and T0.pid='" & sysid & "'"
        dr = SelectLocalSqlUsingDr(SqlCmd, conn)
        If (dr.HasRows) Then
            dr.Read()
            perms = dr(0)
        End If
        dr.Close()
        conn.Close()
        Return perms
    End Function

    Public Function ShowBomData(gv1 As GridView, bomcode As String)
        Dim oBoundField As BoundField
        Dim SqlCmd As String
        Dim connsap As New SqlConnection
        Dim ds As New DataSet
        gv1.AutoGenerateColumns = False
        gv1.AllowPaging = False
        'gv1.Font.Size = FontSize.Small
        gv1.GridLines = GridLines.Both
        gv1.HeaderStyle.HorizontalAlign = HorizontalAlign.Center
        gv1.FooterStyle.HorizontalAlign = HorizontalAlign.Center

        gv1.Columns.Clear()
        oBoundField = New BoundField
        oBoundField.HeaderText = "項次"
        oBoundField.DataField = "icount"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
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
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "用量"
        oBoundField.DataField = "quantity"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F0}"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "倉庫"
        oBoundField.DataField = "whscode"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "庫存"
        oBoundField.DataField = "onhand"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F0}"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "需求"
        oBoundField.DataField = "IsCommited"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F0}"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "供給"
        oBoundField.DataField = "onorder"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F0}"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "它倉庫存"
        oBoundField.DataField = "onhand1"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F0}"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "它倉需求"
        oBoundField.DataField = "IsCommited1"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F0}"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "它倉供給"
        oBoundField.DataField = "onorder1"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F0}"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "安全庫存"
        oBoundField.DataField = "minlevel"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F0}"
        gv1.Columns.Add(oBoundField)
        ds.Reset()
        SqlCmd = "SELECT icount=0,T1.code As Itemcode, T2.ItemName,T1.Quantity, " &
        "T3.WhsCode,T3.OnHand,T3.IsCommited,T3.OnOrder,(T2.OnHand-T3.OnHand) As Onhand1, " &
        "(T2.IsCommited-T3.IsCommited) As IsCommited1,(T2.OnOrder-T3.OnOrder) As OnOrder1, " &
        "T2.MinLevel,T2.InvntItem,T1.Warehouse,IsNull(T2.u_F6,0) FROM OITT T0 " &
        "INNER JOIN ITT1 T1 ON T0.Code = T1.Father " &
        "INNER JOIN OITM T2 ON T2.ItemCode = T1.Code INNER JOIN OITW T3 ON T2.ItemCode=T3.ItemCode " &
        "WHERE T0.[ToWH] = T3.[WhsCode] And T0.Code= '" & bomcode & "'"
        ds = SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
        connsap.Close()
        gv1.DataSource = ds.Tables(0)
        gv1.DataBind()
        If (ds.Tables(0).Rows.Count = 0) Then
            ShowMsg(Me, "無任何資料")
        End If
        Return gv1
    End Function

    Public Function ShowWoData(gv1 As GridView, wo As String)
        Dim oBoundField As BoundField
        Dim SqlCmd As String
        Dim connsap As New SqlConnection
        Dim ds As New DataSet
        gv1.AutoGenerateColumns = False
        gv1.AllowPaging = False
        'gv1.Font.Size = FontSize.Small
        gv1.GridLines = GridLines.Both
        gv1.HeaderStyle.HorizontalAlign = HorizontalAlign.Center
        gv1.FooterStyle.HorizontalAlign = HorizontalAlign.Center

        gv1.Columns.Clear()
        oBoundField = New BoundField
        oBoundField.HeaderText = "項次"
        oBoundField.DataField = "icount"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
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
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "用量"
        oBoundField.DataField = "baseqty"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F0}"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "倉庫"
        oBoundField.DataField = "warehouse"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "庫存"
        oBoundField.DataField = "onhand"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F0}"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "需求"
        oBoundField.DataField = "IsCommited"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F0}"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "供給"
        oBoundField.DataField = "onorder"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F0}"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "它倉庫存"
        oBoundField.DataField = "onhand1"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F0}"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "它倉需求"
        oBoundField.DataField = "IsCommited1"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F0}"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "它倉供給"
        oBoundField.DataField = "onorder1"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F0}"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "計畫"
        oBoundField.DataField = "PlannedQty"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F0}"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "已發"
        oBoundField.DataField = "IssuedQty"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F0}"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "待發"
        oBoundField.DataField = "restQty"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F0}"
        gv1.Columns.Add(oBoundField)

        ds.Reset()
        SqlCmd = "SELECT icount=0,T1.ItemCode, T2.ItemName, T1.BaseQty,T1.warehouse, " &
        "T3.OnHand, T3.IsCommited,T3.OnOrder, T1.PlannedQty, T1.IssuedQty, " &
        "(T2.OnHand-T3.OnHand) As Onhand1, (T2.IsCommited-T3.IsCommited) As IsCommited1,(T2.OnOrder-T3.OnOrder) As OnOrder1, " &
        "T2.InvntItem,(T1.PlannedQty-T1.IssuedQty) As restQty " &
        "FROM OWOR T0 " &
        "INNER JOIN WOR1 T1 ON T0.DocEntry = T1.DocEntry " &
        "INNER JOIN OITM T2 ON T1.ItemCode = T2.ItemCode " &
        "INNER JOIN OITW T3 ON T2.ItemCode = T3.ItemCode " &
        "WHERE T0.warehouse=T3.WhsCode And T0.DocNum = '" & wo & "' " &
        "ORDER BY T1.ItemCode"
        ds = SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
        'Dim dtr() As DataRow
        'dtr = ds.Tables(0).Select("itemcode='0EW1DLP1-4'")
        'If (dtr.Length > 0) Then
        '    MsgBox(dtr(0)("itemname"))
        'Else
        '    MsgBox("none")
        'End If
        connsap.Close()
        gv1.DataSource = ds.Tables(0)
        gv1.DataBind()
        If (ds.Tables(0).Rows.Count = 0) Then
            ShowMsg(Me, "無任何資料")
        End If
        Return gv1
    End Function
    Public Sub SendMail(ToAddress As String, Subject As String, BodyContent As String)
        'Try
        '    Dim myMail As New System.Web.Mail.MailMessage()
        '    myMail.From = "ron@jettech.com.tw"
        '    myMail.To = "ron@jettech.com.tw"
        '    myMail.Subject = "test sub"
        '    myMail.BodyFormat = MailFormat.Html
        '    myMail.Body = "body"
        '    myMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate", "1") 'basic authentication
        '    myMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusername", "ron@jettech.com.tw") 'Set your username here
        '    myMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendpassword", "37512945ron") 'Set your password here
        '    myMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpserverport", 465) 'Set port
        '    myMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpusessl", "true") 'Set Is ssl

        '    System.Web.Mail.SmtpMail.SmtpServer = "smg.jettech.com.tw"
        '    System.Web.Mail.SmtpMail.Send(myMail)
        'Catch
        '    'ex.ToString()
        'End Try
        ToAddress = "ron@jettech.com.tw" 'for 簽核測試用(所有簽核email 都會寄到這)
        Try
            Dim myMail As New MailMessage()
            myMail.From = New MailAddress("jetpmg@jettech.com.tw", "jetpmg") '發送者 	
            myMail.To.Add(New MailAddress(ToAddress))  '收件者
            'myMail.Bcc.Add("456@gmail.com") '隱藏收件者 
            'myMail.CC.Add("789@gmail.com")  '副本 
            myMail.SubjectEncoding = Encoding.UTF8  '主題編碼格式 
            myMail.Subject = Subject '主題 
            myMail.IsBodyHtml = True    'HTML語法(true:開啟false:關閉) 	
            myMail.BodyEncoding = Encoding.UTF8 '內文編碼格式 
            myMail.Body = BodyContent '內文 
            myMail.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure Or DeliveryNotificationOptions.OnSuccess
            'myMail.Attachments.Add(New System.Net.Mail.Attachment("C:\Files\FileA.txt"))  '附件 

            Dim mySmtp As New SmtpClient()  '建立SMTP連線 	
            mySmtp.Credentials = New System.Net.NetworkCredential("jetpmg@jettech.com.tw", "jet#pmg#")  '連線驗證 
            mySmtp.Port = 25   'SMTP Port ==> 無法寄送到大部分mail server , 但jet可以 , 須改用其他方式,如web...
            mySmtp.Host = "smg.jettech.com.tw"  'SMTP主機名 	
            mySmtp.EnableSsl = False '開啟SSL驗證 
            mySmtp.Send(myMail) '發送 	
        Catch
            ShowMsg(Me, "發送失敗")
        End Try
    End Sub
    Public Sub SendMailTxt(ToAddress As String, Subject As String, BodyContent As String)

        Try
            Dim myMail As New MailMessage()
            myMail.From = New MailAddress("jetpmg@jettech.com.tw", "jetpmg") '發送者 	
            myMail.To.Add(New MailAddress(ToAddress))  '收件者
            'myMail.Bcc.Add("456@gmail.com") '隱藏收件者 
            'myMail.CC.Add("789@gmail.com")  '副本 
            myMail.SubjectEncoding = Encoding.UTF8  '主題編碼格式 
            myMail.Subject = Subject '主題 
            myMail.IsBodyHtml = False    'HTML語法(true:開啟false:關閉) 	
            myMail.BodyEncoding = Encoding.UTF8 '內文編碼格式 
            myMail.Body = BodyContent '內文 
            myMail.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure Or DeliveryNotificationOptions.OnSuccess
            'myMail.Attachments.Add(New System.Net.Mail.Attachment("C:\Files\FileA.txt"))  '附件 

            Dim mySmtp As New SmtpClient()  '建立SMTP連線 	
            mySmtp.Credentials = New System.Net.NetworkCredential("jetpmg@jettech.com.tw", "jet#pmg#")  '連線驗證 
            mySmtp.Port = 25   'SMTP Port ==> 無法寄送到大部分mail server , 但jet可以 , 須改用其他方式,如web...
            mySmtp.Host = "smg.jettech.com.tw"  'SMTP主機名 	
            mySmtp.EnableSsl = False '開啟SSL驗證 
            mySmtp.Send(myMail) '發送 	
        Catch
            ShowMsg(Me, "發送失敗")
        End Try
    End Sub
    Public Function GetPostBackControl(page As Page)
        Dim control As Control = Nothing
        Dim controlId As String
        Dim foundControl As Control
        Dim ctl, controlName As String
        controlName = page.Request.Params.Get("__EVENTTARGET")
        If (controlName IsNot Nothing And controlName <> String.Empty) Then 'If (Not String.IsNullOrEmpty(triggerid)) Then
            control = page.FindControl(controlName)
        Else ' 如果是button 之處理
            For Each ctl In page.Request.Form
                If (ctl.EndsWith(".x") Or ctl.EndsWith(".y")) Then
                    controlId = ctl.Substring(0, ctl.Length - 2)
                    foundControl = page.FindControl(controlId)
                Else
                    foundControl = page.FindControl(ctl)
                End If
                If Not (TypeOf foundControl Is Button Or TypeOf foundControl Is ImageButton) Then
                    Continue For
                End If
                control = foundControl
                Exit For
            Next
        End If
        Return control
    End Function
    Function CellSet(text As String, rowspan As Integer, colspan As Integer, FondBold As Boolean, width As Integer, height As Integer, align As String, BColor As Drawing.Color)
        Dim tCell As TableCell
        tCell = New TableCell()
        tCell.BorderWidth = 1
        If (align = "right") Then
            tCell.HorizontalAlign = HorizontalAlign.Right
        ElseIf (align = "center") Then
            tCell.HorizontalAlign = HorizontalAlign.Center
        Else
            tCell.HorizontalAlign = HorizontalAlign.Left
        End If
        'If (color) Then
        tCell.BackColor = BColor
        'End If
        tCell.Wrap = True
        If (text <> "") Then
            tCell.Text = text
        End If
        tCell.ColumnSpan = colspan
        tCell.RowSpan = rowspan
        If (width <> 0) Then
            tCell.Width = width 'stdwidth * colspan * 0.95
        End If
        If (height <> 0) Then
            tCell.Height = height '20 * rowspan
        End If
        tCell.Font.Bold = FondBold
        Return tCell
    End Function

    Function CellSetWithTextBox(rowspan As Integer, colspan As Integer, txtid As String, multiline As Integer, fonesize As Integer, width As Integer, BColor As Drawing.Color, align As String)
        '如果textbox要以預設大小適合 Cell , 則width設為0 , 如果要調大小 , try 看數值多少
        Dim tCell As TableCell
        Dim tTxt As New TextBox
        tCell = New TableCell()
        tCell.Wrap = True
        tCell.BorderWidth = 1
        If (align = "right") Then
            tCell.HorizontalAlign = HorizontalAlign.Right
        ElseIf (align = "center") Then
            tCell.HorizontalAlign = HorizontalAlign.Center
        Else
            tCell.HorizontalAlign = HorizontalAlign.Left
        End If
        tCell.ColumnSpan = colspan
        tCell.RowSpan = rowspan
        tTxt.ID = txtid
        tTxt.BackColor = BColor
        If (fonesize <> 0) Then
            tTxt.Font.Size = fonesize
        End If
        'tTxt.Width = tCell.Width 'stdwidth * colspan * 0.95
        'tTxt.Height = tCell.Height '20 * rowspan
        If (width <> 0) Then
            tTxt.Width = width
        End If
        If (multiline <> 0) Then
            tTxt.TextMode = TextBoxMode.MultiLine
            tTxt.Rows = multiline
        End If
        tCell.Controls.Add(tTxt)
        Return tCell
    End Function
    Function CellSetWithCalenderExtender(rowspan As Integer, colspan As Integer, txtid As String, ceid As String, BColor As Drawing.Color, width As Integer)
        Dim tCell As New TableCell
        Dim ce As New CalendarExtender
        Dim tTxt As New TextBox
        tCell.ColumnSpan = colspan
        tCell.RowSpan = rowspan
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        tTxt.ID = txtid
        If (width <> 0) Then
            tTxt.Width = width
        End If
        tTxt.Height = tCell.Height
        tTxt.BackColor = BColor
        tCell.Controls.Add(tTxt)
        ce.TargetControlID = txtid
        ce.ID = ceid
        ce.Format = "yyyy/MM/dd"
        tCell.Controls.Add(ce)
        Return tCell
    End Function
    Function TextTransToHtmlFormat(displaystr As String)
        Dim str(), kk As String
        kk = ""
        'str = Split(displaystr, Chr(10))
        str = Split(displaystr, vbCrLf)
        For j = 0 To UBound(str)
            kk = kk & str(j) & "<br>"
        Next
        Return kk
    End Function
End Class
