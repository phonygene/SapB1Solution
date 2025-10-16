'1.加入回應加簽單
'2.連續簽核時,加入暫不簽button ==>ok
'3.QC退貨返修管制單
'4.問題反映廠商及回應處理
Imports System.Data
Imports System.Data.SqlClient
Imports AjaxControlToolkit
Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Net
Imports System.Threading
Imports System.Windows
Imports System.Windows.Interop
Imports ExcelDataReader

'Imports O2S.Components.PDFRender4NET
'Imports System.Runtime.InteropServices.DllImportAttribute
'如果有sfid 要特別處理 , 請search 'sfid process' 
'採購單要分AOI/ICT/CNC/外辦 是因這些都是由採購不發出 , 無法像請購是由各部門發出, 可以由部門排除來排除
'sfid=51 (領料單 , 產生簽核list時,在設定表列之人員 ,如果不是全域性簽核者,不做金額判斷) 
Public Class cLsignoff
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public CommSignOff As New CommSignOff
    Public connsap, conn, connsap1, connsap2, connsap10, conn10, connsap11, connsap12 As New SqlConnection
    Public SqlCmd As String
    Public dr, dr1, drsap, dr10, drsap10, dr11, dr12 As SqlDataReader
    Public ds As New DataSet
    Public permssg100 As String
    Public ScriptManager1 As New ScriptManager
    Public RBLType, RBLDpt, RBLCDpt, RBLSpareType As New RadioButtonList
    Public ChkDTDpt As New CheckBoxList
    Public TxtSubject, TxtSapNO, TxtInfo, TxtPrice, TxtAttaDoc As TextBox
    Public DDLDollorUnit As DropDownList
    Public BtnApproval, BtnReject, BtnRecall, BtnSave, BtnSend, BtnDel, BtnCancel, BtnArchieve, BtnNext, BtnLast, BtnPdf, BtnSkip, BtnBeInformed, BtnSuspend As Button
    Public Labelsend, Labeldel, Labelapproval, Labelreject, Labelrecall, Labelcancel, Labelarchieve, LabelNext, LabelLast, LabelCount, LabelPdf, LabelSkip, LabelBeInformed As Label
    Public docnum As Long
    Public docstatus, recall As String
    Public formstatusindex, formtypeindex As Integer
    Public sid, sid_create, agnidG As String
    Public sfid, sftype As Integer
    Public TxtComm As TextBox
    Public actmode, act As String
    Public TxtReason, TxtDept, TxtPerson As TextBox
    Public TxtProblemDescrip, TxtProcessDescrip, TxtVerifyDescrip, TxtProblemNote As TextBox
    Public url, targetPath, targetFile, targetlocalsignofffile As String
    Public info As String
    Public SignOffStatusLabel As Label
    Public signoffalready As Boolean
    Public BtnAssignSend, BtnAssignCancel As Button
    Public DDLSignDefault As DropDownList
    Public localsignoffformdir, localsapuploaddir As String
    Public TxtItemcode, TxtItemname, TxtQty, TxtNote, TxtUnitPrice As TextBox
    Public BtnAction, BtnReset As Button
    Public DDLMethod, DDLMethod1 As DropDownList
    Public signcount, signflowmode As Integer
    Public DDLAttaFile As DropDownList
    Public ChkReturn As CheckBox
    Public ChkUsingAttach As CheckBox
    Public Structure DocAttaL_Data
        Dim docstr As String
        Dim doctype As Integer '0:sub 1:main
    End Structure
    Public DocAttaL(15) As DocAttaL_Data
    Public inchargeindex As Integer
    Public traceindex As Integer
    Public maindocnum As Long
    Public PreAttaDoc, inchargeid, fromasp As String
    Public indexpage As Integer
    Public gseq As Integer
    Public signpersonmaxrow As Integer
    'Public MainAttachFile As Boolean
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.NoCache)
        'HttpContext.Current.Response.Cache.SetNoServerCaching()
        'HttpContext.Current.Response.Cache.SetNoStore()
        Dim s_name As String
        Dim Now_time As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        'MainAttachFile = False
        signcount = 0
        gseq = 1
        signpersonmaxrow = 17
        info = ""
        sid = Session("s_id")
        s_name = Session("s_name")

        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If

        signflowmode = 0
        Page.Form.Controls.Add(ScriptManager1)
        url = Application("http")
        localsapuploaddir = Application("localdir")
        act = Request.QueryString("act")
        fromasp = Request.QueryString("fromasp")
        If (Not IsPostBack) Then
            indexpage = Request.QueryString("indexpage")
            agnidG = Request.QueryString("agnid")
            If (agnidG <> "") Then
                If (CommSignOff.AgencySet(sid) = "") Then
                    Response.Redirect("~\invalid.aspx?info=代理簽核授權已結束")
                End If
            End If
            docstatus = Request.QueryString("status")
            sfid = Request.QueryString("sfid")
            formtypeindex = Request.QueryString("formtypeindex")
            actmode = Request.QueryString("actmode")
            signflowmode = Request.QueryString("signflowmode")
            maindocnum = Request.QueryString("maindocnum")
            If (fromasp = "signofftodo") Then
                inchargeindex = Request.QueryString("inchargeindex")
                traceindex = Request.QueryString("traceindex")
                inchargeid = Request.QueryString("inchargeid")
            Else
                formstatusindex = Request.QueryString("formstatusindex")
            End If
            SqlCmd = "select T0.sftype from [dbo].[@XSFTT] T0 " & 'XSFTT 簽核表單種類
                "where T0.sfid=" & sfid
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                dr.Read()
                sftype = dr(0)
            End If
            dr.Close()
            connsap.Close()
            CT.Visible = False
            If (docstatus = "A") Then
                sid_create = sid
            Else
                docnum = Request.QueryString("docnum")
                'get 此單據送審人 , 以獲得所存附件路徑
                SqlCmd = "Select sid,status from [dbo].[@XASCH] where docnum=" & docnum
                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                If (dr.HasRows) Then
                    dr.Read()
                    sid_create = dr(0)
                    docstatus = dr(1)
                End If
                dr.Close()
                connsap.Close()
            End If
            ViewState("pageindex") = indexpage
            ViewState("docstatus") = docstatus
            ViewState("sfid") = sfid
            ViewState("formstatusindex") = formstatusindex
            ViewState("formtypeindex") = formtypeindex
            ViewState("docnum") = docnum
            ViewState("sid_create") = sid_create
            ViewState("actmode") = actmode
            ViewState("sftype") = sftype
            ViewState("agnid") = agnidG
            ViewState("signflowmode") = signflowmode
            ViewState("maindocnum") = maindocnum
            If (fromasp = "signofftodo") Then
                ViewState("inchargeindex") = inchargeindex
                ViewState("traceindex") = traceindex
                ViewState("inchargeid") = inchargeid
            End If
            If (act = "save") Then
                CommUtil.ShowMsg(Me, "資料完整已存檔 ,可執行送審")
            ElseIf (act = "del") Then
                CommUtil.ShowMsg(Me, "已刪除單據")
            ElseIf (act = "fileupanddel") Then
                CommUtil.ShowMsg(Me, "附檔上傳成功,原檔案已刪除")
            ElseIf (act = "fileup") Then
                CommUtil.ShowMsg(Me, "附檔上傳成功")
            ElseIf (act = "filenoassign") Then
                CommUtil.ShowMsg(Me, "未指定附檔")
            ElseIf (act = "filedel") Then
                CommUtil.ShowMsg(Me, "附檔已刪除")
            ElseIf (act = "skipok") Then
                CommUtil.ShowMsg(Me, "已跳過簽核,進入下一關")
            ElseIf (act = "create") Then
                CommUtil.ShowMsg(Me, "已建立新簽核單")
            ElseIf (act = "notsave" And sfid <> 23) Then
                CommUtil.ShowMsg(Me, "已儲存,但還有欄位未填入或沒附檔(或料件)")
            ElseIf (act = "notsaveattach") Then
                CommUtil.ShowMsg(Me, "已儲存,但還沒上傳簽核主檔")
            ElseIf (act = "material_del") Then
                SqlCmd = "delete from [dbo].[@XSMLS] where num=" & Request.QueryString("num")
                CommUtil.SqlSapExecute("del", SqlCmd, connsap)
                connsap.Close()
            End If
            If (actmode = "signoff" Or actmode = "signoff_login") Then '連續審核mode
                SqlCmd = "Select sfid from [dbo].[@XASCH] where docnum=" & docnum '因怕sfid有時要調整 , 已寄的通知sfid就會對不上,故依此為主
                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                If (dr.HasRows) Then
                    dr.Read()
                    sfid = dr(0)
                End If
                dr.Close()
                connsap.Close()
                Session("startindex") = GetSinoffList() 'begin 0
                Session("ds") = ds
                If (actmode = "signoff") Then
                    actmode = "recycle"
                    ViewState("actmode") = "recycle"
                Else
                    actmode = "recycle_login"
                    ViewState("actmode") = "recycle_login"
                End If
            ElseIf (actmode = "recycle" Or actmode = "recycle_login") Then '開始連續審核
                ds = Session("ds")
                ViewState("actmode") = actmode
            End If
        Else
            indexpage = ViewState("indexpage")
            sid_create = ViewState("sid_create")
            docnum = ViewState("docnum")
            docstatus = ViewState("docstatus")
            formstatusindex = ViewState("formstatusindex")
            formtypeindex = ViewState("formtypeindex")
            sfid = ViewState("sfid")
            actmode = ViewState("actmode")
            sftype = ViewState("sftype")
            agnidG = ViewState("agnid")
            If (actmode = "recycle" Or actmode = "recycle_login") Then '開始連續審核
                ds = Session("ds")
            End If
            signflowmode = ViewState("signflowmode")
            maindocnum = ViewState("maindocnum")
            If (fromasp = "signofftodo") Then
                inchargeindex = ViewState("inchargeindex")
                traceindex = ViewState("traceindex")
                inchargeid = ViewState("inchargeid")
            End If
        End If
        'CCPersonProcess() '知悉人員設定==>改為不自動 , 按已知悉button 才設定(因要將知悉人納入連續簽核表列 , 故需一button來進入下一張表單)
        CreateFormInfo() 'headT 嵌入物件 , 要放在FTCreate前 , 因有物件在此需設定
        FT0Create() ''futher maybe use
        FT_0.Visible = False 'futher maybe use
        FTCreate()
        Page.Form.Controls.Add(ScriptManager1)
        CreateSignFlowHistoryField() 'common
        CreateCommentField() 'common
        CreateSignFlowPerson() 'set up 自定簽核人輸入==>CT table Create
        FT1Create()
        If (Not IsPostBack) Then
            If (docstatus <> "A") Then
                PutDataToFormInfo() ' 將此放在ContenTCreate上面(主要是其中之showmaterial) , 否則在showmaterial中 price 更新txtprice後又會被此原始資料蓋掉
            Else
                FillHeadData()
            End If
        End If
        If (docstatus <> "A") Then
            GenAttaFileList()
        Else
            If (sfid = 100 Or sfid = 101) Then
                GenMainAttaFile_sfid100_101()
            End If
        End If
        localsignoffformdir = Application("localdir") & "SignOffsFormFiles\" & sid_create & "\" & sfid & "\"
        targetPath = HttpContext.Current.Server.MapPath("~/") & "AttachFile\SignOffsFormFiles\" & sid_create & "\" & sfid & "\"

        targetFile = targetPath & DDLAttaFile.SelectedValue
        targetlocalsignofffile = localsignoffformdir & DDLAttaFile.SelectedValue
        If (DDLAttaFile.SelectedIndex <> 0 And DDLAttaFile.SelectedIndex <> -1) Then
            Dim httpfile As String
            Dim siddir As String
            Dim sfidnum As Integer
            Dim str() As String
            str = Split(DDLAttaFile.SelectedValue, "_")
            SqlCmd = "Select sid,sfid from [dbo].[@XASCH] where docnum=" & CLng(str(0))
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            siddir = dr(0)
            sfidnum = dr(1)
            dr.Close()
            connsap.Close()
            httpfile = url & "AttachFile/SignOffsFormFiles/" & siddir & "/" & sfidnum & "/" & DDLAttaFile.SelectedValue
            iframeContent.Attributes.Remove("src")
            iframeContent.Attributes.Add("src", httpfile)
        Else
            If (DDLAttaFile.Items.Count = 1) Then
                iframeContent.Visible = False
                FT_1.Visible = False
            End If
        End If
        If (sfid = 51 Or sfid = 50 Or sfid = 49 Or sfid = 100 Or sfid = 23 Or sfid = 24) Then
            AddTCreate() '料件增加修改功能欄 for sfid 51,50,49
        End If
        ContentTCreate() '料件List Table
        If (docstatus <> "A" And docstatus <> "E" And docstatus <> "D" And docstatus <> "R" And docstatus <> "B") Then
            ContentT.Enabled = False
        End If
        If (docstatus <> "F" And docstatus <> "T") Then
            SqlCmd = "select seq from  [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum & " order by seq"
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                Do While (dr.Read())
                    CreateSignOffFlowField(dr(0))
                Loop
            End If
            dr.Close()
            connsap.Close()
            PutDataToSignOffFlow()
        End If
        If (HeadT.Rows(0).Cells(1).Text = "") Then
            FileUpLoadObjectDisabled()
        Else
            FileUpLoadObjectEnabled()
        End If
        If (CommSignOff.IsSelfForm(sfid) = 0) Then 'sfid process
            AddT.Visible = False
            ContentT.Visible = False
            FormLogoTitleT.Visible = False
        Else
            If (docstatus <> "F" And docstatus <> "T") Then
                'iframeContent.Visible = False
                'FileUpLoadObjectDisabled()
                'If (sfid = 51 Or sfid = 50 Or sfid = 49) Then
                If (HeadT.Rows(0).Cells(1).Text = "") Then
                    AddT.Enabled = False
                Else
                    If (docstatus = "E" Or docstatus = "D" Or docstatus = "B" Or docstatus = "R") Then
                        AddT.Visible = True '若遇CT.Visible , 則最底行再設為 False , 就不再這做判斷
                    Else
                        AddT.Visible = False
                    End If
                End If
                'End If
            Else
                'If (sfid = 51 Or sfid = 50 Or sfid = 49) Then
                AddT.Visible = False
                'End If
                If (maindocnum = 0) Then '若是由signofftodo.aspx而來 , 則以下不設為False
                    ContentT.Visible = False
                    FormLogoTitleT.Visible = False
                End If
            End If
            TxtPrice.Enabled = False
            DDLDollorUnit.SelectedValue = "NTD"
            DDLDollorUnit.Enabled = False
            If (sfid = 16 Or sfid = 12) Then
                'AddT.Visible = False
                DDLDollorUnit.SelectedIndex = 0
                TxtSubject.Enabled = False
                If (docstatus = "A") Then
                    TxtSubject.Text = "NA"
                End If
            Else
                If (docstatus = "A") Then
                    AddT.Enabled = False
                End If
            End If
        End If
        PreAttaDoc = TxtAttaDoc.Text
        If (CT.Visible = True) Then
            AddT.Visible = False
        End If
    End Sub
    Sub CCPersonProcess()
        Dim signdate As String
        signdate = Format(Now(), "yyyy/MM/dd HH:mm:ss")
        SqlCmd = "select count(*) from  [dbo].[@XSPWT] T0 " &
                "where signprop=2 And docentry=" & docnum & " And uid='" & Session("s_id") & "' and T0.status=1"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap2)
        dr.Read()
        If (dr(0) <> 0) Then
            '把status update為104
            SqlCmd = "update [dbo].[@XSPWT] set status=104,signdate='" & signdate & "' where signprop=2 and docentry=" & docnum & " and uid='" & Session("s_id") & "'"
            CommUtil.SqlSapExecute("upd", SqlCmd, connsap1)
            connsap1.Close()
            '寫入History
            '在@XSPHT(簽核歷史Table)記錄此核准資料
            RecordSignFlowHistoty("", "已知悉")
        End If
        dr.Close()
        connsap2.Close()
    End Sub
    Sub GenAttaFileList()
        DDLAttaFile.Items.Clear()
        DDLAttaFile.Items.Add("請選擇欲顯示之附加檔案")
        Dim di As DirectoryInfo
        Dim attachsel, str() As String
        Dim maindoc As Boolean
        maindoc = False
        'Dim docarr(10) As String
        Dim i, j As Integer
        i = 0
        j = 0
        attachsel = ""
        DocAttaL(0).docstr = "end"
        If (sfid = 100 Or sfid = 101) Then '分現處理的是加簽還是ㄧ般  (加簽且是本單放第一個 , 後陸續放加簽 , 最後一個放母單號)
            If (TxtAttaDoc.Text <> "" And TxtAttaDoc.Text <> "NA") Then
                'MsgBox(TxtAttaDoc.Text)
                maindoc = True
                DocAttaL(i).docstr = CStr(docnum)
                DocAttaL(i).doctype = 0
                i = i + 1
                SqlCmd = "Select attadoc from [dbo].[@XASCH] where docnum=" & CLng(TxtAttaDoc.Text)
                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                dr.Read()
                str = Split(dr(0), "_")
                dr.Close()
                connsap.Close()
                If (str(0) = "NA") Then
                    DocAttaL(i).docstr = TxtAttaDoc.Text
                    DocAttaL(i).doctype = 1
                    i = i + 1
                    DocAttaL(i).docstr = "end"
                Else
                    For j = 0 To UBound(str) ''UBOUND: 查詢陣列有幾個 1:表示2個
                        If (CLng(str(j)) <> docnum) Then
                            DocAttaL(i).docstr = str(j)
                            DocAttaL(i).doctype = 0
                            i = i + 1
                        End If
                    Next
                    DocAttaL(i).docstr = TxtAttaDoc.Text
                    DocAttaL(i).doctype = 1
                    i = i + 1
                    DocAttaL(i).docstr = "end"
                End If
            Else
                If (docstatus <> "A") Then
                    DocAttaL(i).docstr = CStr(docnum)
                    DocAttaL(i).doctype = 0
                    i = i + 1
                    DocAttaL(i).docstr = "end"
                End If
            End If
        Else
            If (docstatus <> "A") Then
                'SqlCmd = "Select attadoc from [dbo].[@XASCH] where docnum=" & docnum
                'dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                'dr.Read()
                str = Split(TxtAttaDoc.Text, "_")
                'dr.Close()
                'connsap.Close()
                If (str(0) = "NA") Then
                    DocAttaL(i).docstr = CStr(docnum)
                    DocAttaL(i).doctype = 1
                    i = i + 1
                    DocAttaL(i).docstr = "end"
                Else
                    For j = 0 To UBound(str) ''UBOUND: 查詢陣列有幾個 1:表示2個
                        DocAttaL(i).docstr = str(j)
                        DocAttaL(i).doctype = 0
                        i = i + 1
                    Next
                    DocAttaL(i).docstr = CStr(docnum)
                    DocAttaL(i).doctype = 1
                    i = i + 1
                    DocAttaL(i).docstr = "end"
                End If
                maindoc = True
            End If
        End If
        i = 0
        j = 0
        Dim siddir As String
        Dim sfidnum As Integer
        Dim formstatus As String
        Dim docnumloop As Long
        Do While (DocAttaL(j).docstr <> "end")
            'MsgBox(DocAttaL(j).docstr)
            SqlCmd = "Select sid,sfid,status,docnum from [dbo].[@XASCH] where docnum=" & CLng(DocAttaL(j).docstr)
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            siddir = dr(0)
            sfidnum = dr(1)
            formstatus = dr(2)
            docnumloop = dr(3)
            dr.Close()
            connsap.Close()
            If (System.IO.Directory.Exists(HttpContext.Current.Server.MapPath("~/") & "AttachFile\SignOffsFormFiles\" & siddir & "\" & sfidnum & "\")) Then
                di = New DirectoryInfo(HttpContext.Current.Server.MapPath("~/") & "AttachFile\SignOffsFormFiles\" & siddir & "\" & sfidnum & "\")
                Dim fi As FileInfo()
                If (sfidnum <> 100) Then
                    fi = di.GetFiles(docnumloop & "_簽核*")
                Else
                    fi = di.GetFiles(docnumloop & "_加簽核*")
                End If
                For i = 0 To fi.Length - 1
                    DDLAttaFile.Items.Add(fi(i).Name)
                    If (Request.QueryString("attachsel") = fi(i).Name) Then '如果是refresh 並且想要顯示refresh之前檔案
                        attachsel = Request.QueryString("attachsel")
                    Else
                        If (maindoc And DocAttaL(j).doctype = 1) Then
                            If (formstatus <> "F" And formstatus <> "T") Then
                                If (CommSignOff.IsSelfForm(sfid) = 0) Then '自行Import 之主檔 , 且未產生stamped
                                    If (InStr(fi(i).Name, "主") <> 0 And (InStr(fi(i).Name, "Stamped") = 0)) Then
                                        attachsel = fi(i).Name
                                    End If
                                Else '是內建表格 , 故簽核主檔是自行產生,故此時尚未有主檔 , 故show附檔(如果有pdf的話)
                                    If (InStr(fi(i).Name, "附") <> 0 And (InStr(fi(i).Name, ".pdf") <> 0)) Then
                                        attachsel = fi(i).Name
                                    End If
                                End If
                            Else '已歸檔並產生Stamped file
                                If (InStr(fi(i).Name, "主") <> 0 And InStr(fi(i).Name, "Stamped") <> 0) Then
                                    attachsel = fi(i).Name
                                End If
                            End If
                        Else
                            If (j = 0) Then
                                attachsel = fi(i).Name
                            End If
                        End If
                    End If
                Next
                'DDLAttaFile.SelectedValue = attachsel
            End If
            j = j + 1
        Loop
        If (attachsel <> "") Then
            DDLAttaFile.SelectedValue = attachsel
        End If
    End Sub
    Sub GenMainAttaFile_sfid100_101()
        DDLAttaFile.Items.Clear()
        DDLAttaFile.Items.Add("請選擇欲顯示之附加檔案")
        Dim di As DirectoryInfo
        Dim attachsel, str() As String
        Dim maindoc As Boolean
        maindoc = False
        'Dim docarr(10) As String
        Dim i, j As Integer
        i = 0
        j = 0
        attachsel = ""
        DocAttaL(0).docstr = "end"
        If (sfid = 100 Or sfid = 101) Then '加簽且是本單放第一個 , 後陸續放加簽 , 最後一個放母單號
            If (TxtAttaDoc.Text <> "" And TxtAttaDoc.Text <> "NA") Then
                'MsgBox(TxtAttaDoc.Text)
                maindoc = True
                SqlCmd = "Select attadoc from [dbo].[@XASCH] where docnum=" & CLng(TxtAttaDoc.Text)
                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                dr.Read()
                str = Split(dr(0), "_")
                dr.Close()
                connsap.Close()
                If (str(0) = "NA") Then
                    DocAttaL(i).docstr = TxtAttaDoc.Text
                    DocAttaL(i).doctype = 1
                    i = i + 1
                    DocAttaL(i).docstr = "end"
                Else
                    For j = 0 To UBound(str) ''UBOUND: 查詢陣列有幾個 1:表示2個
                        If (CLng(str(j)) <> docnum) Then
                            DocAttaL(i).docstr = str(j)
                            DocAttaL(i).doctype = 0
                            i = i + 1
                        End If
                    Next
                    DocAttaL(i).docstr = TxtAttaDoc.Text
                    DocAttaL(i).doctype = 1
                    i = i + 1
                    DocAttaL(i).docstr = "end"
                End If
            Else
                '尚未給與母單號 , 故無法先show母單檔案
            End If
        End If
        i = 0
        j = 0
        Dim siddir As String
        Dim sfidnum As Integer
        Dim formstatus As String
        Dim docnumloop As Long
        Do While (DocAttaL(j).docstr <> "end")
            'MsgBox(DocAttaL(j).docstr)
            SqlCmd = "Select sid,sfid,status,docnum from [dbo].[@XASCH] where docnum=" & CLng(DocAttaL(j).docstr)
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            siddir = dr(0)
            sfidnum = dr(1)
            formstatus = dr(2)
            docnumloop = dr(3)
            dr.Close()
            connsap.Close()
            If (System.IO.Directory.Exists(HttpContext.Current.Server.MapPath("~/") & "AttachFile\SignOffsFormFiles\" & siddir & "\" & sfidnum & "\")) Then
                di = New DirectoryInfo(HttpContext.Current.Server.MapPath("~/") & "AttachFile\SignOffsFormFiles\" & siddir & "\" & sfidnum & "\")
                Dim fi As FileInfo()
                If (sfidnum <> 100) Then
                    fi = di.GetFiles(docnumloop & "_簽核*")
                Else
                    fi = di.GetFiles(docnumloop & "_加簽核*")
                End If
                For i = 0 To fi.Length - 1
                    DDLAttaFile.Items.Add(fi(i).Name)
                    If (Request.QueryString("attachsel") = fi(i).Name) Then '如果是refresh 並且想要顯示refresh之前檔案
                        attachsel = Request.QueryString("attachsel")
                    Else
                        If (maindoc And DocAttaL(j).doctype = 1) Then
                            If (formstatus <> "F" And formstatus <> "T") Then
                                If (CommSignOff.IsSelfForm(sfid) = 0) Then '自行Import 之主檔 , 且未產生stamped
                                    If (InStr(fi(i).Name, "主") <> 0 And (InStr(fi(i).Name, "Stamped") = 0)) Then
                                        attachsel = fi(i).Name
                                    End If
                                Else '是內建表格 , 故簽核主檔是自行產生,故此時尚未有主檔 , 故show附檔(如果有pdf的話)
                                    If (InStr(fi(i).Name, "附") <> 0 And (InStr(fi(i).Name, ".pdf") <> 0)) Then
                                        attachsel = fi(i).Name
                                    End If
                                End If
                            Else '已歸檔並產生Stamped file
                                If (InStr(fi(i).Name, "主") <> 0 And InStr(fi(i).Name, "Stamped") <> 0) Then
                                    attachsel = fi(i).Name
                                End If
                            End If
                        Else
                            If (j = 0) Then
                                attachsel = fi(i).Name
                            End If
                        End If
                    End If
                Next
                'DDLAttaFile.SelectedValue = attachsel
            End If
            j = j + 1
        Loop
        If (attachsel <> "") Then
            DDLAttaFile.SelectedValue = attachsel
        End If
    End Sub
    Sub FTCreate()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim Hyper As HyperLink
        Dim Labelx, LabelFileu, LabelUpfile As Label
        Dim sign_status, nowseq, innerloop As Integer
        Dim lastsignoff As Boolean

        lastsignoff = False
        sign_status = 0
        Dim signofftype, cc As Integer
        signofftype = 0
        SqlCmd = "Select count(*) from [dbo].[@XSPWT] where docentry=" & docnum & " and uid='" & Session("s_id") & "'"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        dr.Read()
        cc = dr(0)
        dr.Close()
        connsap.Close()
        SqlCmd = "Select status,signprop,seq,innerloop from [dbo].[@XSPWT] where docentry=" & docnum & " and uid='" & Session("s_id") & "' order by seq"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        dr.Read()
        If (cc <> 0) Then
            nowseq = dr(2)
            innerloop = dr(3)
        End If
        If (cc = 1) Then '所有在簽核表中沒有重覆id
            sign_status = dr(0)
            signofftype = dr(1)
        ElseIf (cc = 0) Then

        Else '審核者及歸檔者是同一個
            If (docstatus = "D" Or docstatus = "B" Or docstatus = "R") Then '是審核者
                If (dr(2) <> 1) Then '非審核者 , 故需再read得到審核者
                    dr.Read()
                End If
            ElseIf (docstatus = "O") Then
                'nothing
            ElseIf (docstatus = "F" Or docstatus = "T") Then '是歸檔者
                'If (dr(2) = 1) Then '非歸檔者 , 故需再read得到歸檔者
                dr.Read()
                'End If
            End If
            sign_status = dr(0)
            signofftype = dr(1)
        End If
        dr.Close()
        connsap.Close()
        If (docstatus = "D") Then
            sign_status = 1
        End If
        If (cc <> 0) Then
            SqlCmd = "Select max(seq) from [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            If (dr(0) = nowseq) Then
                lastsignoff = True
            End If
            dr.Close()
            connsap.Close()
        End If
        tRow = New TableRow()
        tCell = New TableCell()
        'tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Left
        Hyper = New HyperLink
        Hyper.ID = "hyper_back"
        Hyper.Text = "回前頁"
        If (fromasp = "signofftodo") Then
            Hyper.NavigateUrl = "~/signoff/signofftodo.aspx?smid=sg&smode=6&inchargeindex=" & inchargeindex & "&traceindex=" & traceindex &
                            "&formtypeindex=" & formtypeindex & "&indexpage=" & indexpage & "&inchargeid=" & inchargeid
        Else
            If (act = "skip" Or act = "skipok" Or act = "frommanage") Then
                Hyper.NavigateUrl = "~/signoff/signoffvip.aspx?smid=sg&smode=5&signflowmode=" & signflowmode & "&formstatusindex=" & formstatusindex &
                                "&formtypeindex=" & formtypeindex & "&indexpage=" & indexpage
            Else
                Hyper.NavigateUrl = "~/signoff/signoff.aspx?smid=sg&smode=1&signflowmode=" & signflowmode & "&formstatusindex=" & formstatusindex &
                                "&formtypeindex=" & formtypeindex & "&indexpage=" & indexpage
            End If
        End If
        Hyper.Font.Underline = False
        tCell.Controls.Add(Hyper)
        If (actmode = "signoff" Or actmode = "recycle" Or actmode = "single_signoff") Then
            Hyper.Visible = False
        Else
            Hyper.Visible = True
        End If

        Labelx = New Label()
        Labelx.ID = "label_subfile"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        Dim ChkFileType As New CheckBox
        ChkFileType.ID = "chk_subfile"
        ChkFileType.Text = "非主檔"
        'ChkFileType.AutoPostBack = True
        'AddHandler ChkFileType.CheckedChanged, AddressOf ChkFileType_CheckedChanged
        tCell.Controls.Add(ChkFileType)

        LabelFileu = New Label()
        LabelFileu.ID = "label_fileul"
        LabelFileu.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(LabelFileu)
        Dim FileUL As New FileUpload()
        FileUL.ID = "fileul_m"
        tCell.Controls.Add(FileUL)
        'If (docstatus = "A" Or docstatus = "E" Or docstatus = "D" Or ((docstatus = "R" Or docstatus = "B") And formstatusindex = 0 And sign_status = 1)) Then
        '    LabelFileu.Visible = True
        '    FileUL.Visible = True
        'Else
        '    LabelFileu.Visible = False
        '    FileUL.Visible = False
        'End If

        Dim ChkDel As New CheckBox
        ChkDel.ID = "chk_del_m"
        ChkDel.Text = "刪檔"
        ChkDel.AutoPostBack = True
        AddHandler ChkDel.CheckedChanged, AddressOf ChkDel_CheckedChanged
        tCell.Controls.Add(ChkDel)

        LabelUpfile = New Label()
        LabelUpfile.ID = "label_upfile"
        LabelUpfile.Text = "&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(LabelUpfile)
        Dim BtnFileAct As New Button
        BtnFileAct.ID = "btn_fileact_m"
        BtnFileAct.Text = "上傳"
        AddHandler BtnFileAct.Click, AddressOf BtnFileAct_Click
        tCell.Controls.Add(BtnFileAct)
        'If (docstatus = "D" Or docstatus = "A" Or docstatus = "E" Or ((docstatus = "R" Or docstatus = "B") And formstatusindex = 0 And sign_status = 1)) Then
        '    LabelUpfile.Visible = True
        '    ChkDel.Visible = True
        '    BtnFileAct.Visible = True
        '    ChkFileType.Visible = True
        '    If (CommSignOff.IsSelfForm(sfid) = 1) Then
        '        ChkFileType.Checked = True
        '        ChkFileType.Enabled = False
        '    Else
        '        ChkFileType.Checked = False
        '        ChkFileType.Enabled = True
        '    End If
        'Else
        '    LabelUpfile.Visible = False
        '    ChkDel.Visible = False
        '    BtnFileAct.Visible = False
        '    ChkFileType.Visible = False
        'End If

        tRow.Cells.Add(tCell)

        Labelx = New Label()
        Labelx.ID = "label_ddlattafile"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        DDLAttaFile = New DropDownList
        DDLAttaFile.ID = "ddl_attafile"
        DDLAttaFile.Width = 240
        AddHandler DDLAttaFile.SelectedIndexChanged, AddressOf DDLAttaFile_SelectedIndexChanged
        DDLAttaFile.AutoPostBack = True
        DDLAttaFile.Items.Clear()
        DDLAttaFile.Items.Add("請選擇欲顯示之附加檔案")
        DDLAttaFile.SelectedIndex = 0 '如果無上述2個statement , 此值會為-1 (就算指定為0也是會為-1)
        tCell.Controls.Add(DDLAttaFile)
        'Dim di As DirectoryInfo
        'Dim attachsel As String
        'attachsel = ""
        'If (System.IO.Directory.Exists(HttpContext.Current.Server.MapPath("~/") & "AttachFile\SignOffsFormFiles\" & sid_create & "\" & sfid & "\")) Then
        '    di = New DirectoryInfo(HttpContext.Current.Server.MapPath("~/") & "AttachFile\SignOffsFormFiles\" & sid_create & "\" & sfid & "\")
        '    Dim fi As FileInfo() = di.GetFiles(docnum & "_簽核*")
        '    For i = 0 To fi.Length - 1
        '        DDLAttaFile.Items.Add(fi(i).Name)
        '        If (Request.QueryString("attachsel") = fi(i).Name) Then '如果是refresh 並且想要顯示refresh之前檔案
        '            attachsel = Request.QueryString("attachsel")
        '        Else
        '            If (docstatus <> "F" And docstatus <> "T") Then
        '                If (CommSignOff.IsSelfForm(sfid) = 0) Then '自行Import 之主檔 , 且未產生stamped
        '                    If (InStr(fi(i).Name, "主") <> 0 And (InStr(fi(i).Name, "Stamped") = 0)) Then
        '                        attachsel = fi(i).Name
        '                        'MainAttachFile = True
        '                    End If
        '                Else '是內建表格 , 故簽核主檔是自行產生,故此時尚未有主檔 , 故show附檔(如果有pdf的話)
        '                    If (InStr(fi(i).Name, "附") <> 0 And (InStr(fi(i).Name, ".pdf") <> 0)) Then
        '                        attachsel = fi(i).Name
        '                        'MainAttachFile = True
        '                    End If
        '                End If
        '            Else '已歸檔並產生Stamped file
        '                If (InStr(fi(i).Name, "主") <> 0 And InStr(fi(i).Name, "Stamped") <> 0) Then
        '                    attachsel = fi(i).Name
        '                    'MainAttachFile = True
        '                End If
        '            End If
        '        End If
        '    Next
        '    DDLAttaFile.SelectedValue = attachsel
        'End If
        tRow.Controls.Add(tCell)
        If (DDLAttaFile.SelectedIndex = 0) Then
            ChkDel.Visible = False
        End If

        tCell = New TableCell()
        'tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Right
        BtnSave = New Button
        BtnSave.ID = "btn_save"
        BtnSave.Width = 60
        If (docstatus = "A") Then
            BtnSave.Text = "建單"
        Else
            BtnSave.Text = "儲存"
        End If
        AddHandler BtnSave.Click, AddressOf BtnSave_Click
        tCell.Controls.Add(BtnSave)
        Labelsend = New Label
        Labelsend.ID = "label_send"
        Labelsend.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelsend)
        BtnSend = New Button
        BtnSend.ID = "btn_send"
        BtnSend.Width = 60
        BtnSend.Text = "送審"
        'act = "edit"
        AddHandler BtnSend.Click, AddressOf BtnSend_Click
        tCell.Controls.Add(BtnSend)

        Labeldel = New Label
        Labeldel.ID = "label_del"
        Labeldel.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labeldel)
        BtnDel = New Button
        BtnDel.ID = "btn_del"
        BtnDel.Width = 60
        BtnDel.Text = "刪除"
        AddHandler BtnDel.Click, AddressOf BtnDel_Click
        tCell.Controls.Add(BtnDel)

        Labelapproval = New Label
        Labelapproval.ID = "label_approval"
        Labelapproval.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelapproval)
        BtnApproval = New Button
        BtnApproval.ID = "btn_approval"
        BtnApproval.Width = 60
        BtnApproval.Text = "核准"
        AddHandler BtnApproval.Click, AddressOf BtnApproval_Click
        tCell.Controls.Add(BtnApproval)

        Labelreject = New Label
        Labelreject.ID = "label_reject"
        Labelreject.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelreject)
        BtnReject = New Button
        BtnReject.ID = "btn_reject"
        BtnReject.Width = 60
        If (cc <> 0) Then
            If (innerloop = 1 Or lastsignoff = True) Then
                BtnReject.Text = "駁回"
            Else
                BtnReject.Text = "反對"
            End If
        Else
            BtnReject.Text = "駁回"
        End If
        AddHandler BtnReject.Click, AddressOf BtnReject_Click
        tCell.Controls.Add(BtnReject)

        Labelx = New Label
        Labelx.ID = "label_suspend"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        BtnSuspend = New Button
        BtnSuspend.ID = "btn_suspend"
        BtnSuspend.Width = 70
        BtnSuspend.Text = "暫不簽核"
        AddHandler BtnSuspend.Click, AddressOf BtnNext_Click
        tCell.Controls.Add(BtnSuspend)

        'Labelx = New Label()
        'Labelx.ID = "label_return"
        'Labelx.Text = "&nbsp"
        'tCell.Controls.Add(Labelx)
        ChkReturn = New CheckBox
        ChkReturn.ID = "chk_return_m"
        ChkReturn.Text = "退回審核人"
        AddHandler ChkReturn.CheckedChanged, AddressOf ChkReturn_CheckedChanged
        tCell.Controls.Add(ChkReturn)
        If (docstatus = "F" Or docstatus = "T") Then
            ChkReturn.Visible = False
        Else
            If (cc <> 0) Then
                If (innerloop = 1 Or lastsignoff = True) Then
                    ChkReturn.Visible = False
                Else
                    ChkReturn.Visible = True
                End If
            Else
                ChkReturn.Visible = False
            End If
        End If
        Labelrecall = New Label
        Labelrecall.ID = "label_recall"
        Labelrecall.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelrecall)
        BtnRecall = New Button
        BtnRecall.ID = "btn_recall"
        BtnRecall.Width = 60
        BtnRecall.Text = "抽回"
        AddHandler BtnRecall.Click, AddressOf BtnRecall_Click
        tCell.Controls.Add(BtnRecall)

        Labelcancel = New Label
        Labelcancel.ID = "label_cancel"
        Labelcancel.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelcancel)
        BtnCancel = New Button
        BtnCancel.ID = "btn_cancel"
        BtnCancel.Width = 60
        BtnCancel.Text = "作廢"
        AddHandler BtnCancel.Click, AddressOf BtnCancel_Click
        tCell.Controls.Add(BtnCancel)

        Labelarchieve = New Label
        Labelarchieve.ID = "label_archieve"
        Labelarchieve.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelarchieve)
        BtnArchieve = New Button
        BtnArchieve.ID = "btn_archieve"
        BtnArchieve.Width = 60
        BtnArchieve.Text = "歸檔"
        AddHandler BtnArchieve.Click, AddressOf BtnArchieve_Click
        tCell.Controls.Add(BtnArchieve)

        LabelBeInformed = New Label
        LabelBeInformed.ID = "label_BeInformed"
        LabelBeInformed.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(LabelBeInformed)
        BtnBeInformed = New Button
        BtnBeInformed.ID = "btn_BeInformed"
        BtnBeInformed.Width = 60
        BtnBeInformed.Text = "已知悉"
        AddHandler BtnBeInformed.Click, AddressOf BtnBeInformed_Click
        tCell.Controls.Add(BtnBeInformed)

        LabelSkip = New Label
        LabelSkip.ID = "label_skip"
        LabelSkip.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(LabelSkip)
        BtnSkip = New Button
        BtnSkip.ID = "btn_skip"
        BtnSkip.Width = 70
        BtnSkip.Text = "跳過簽核"
        AddHandler BtnSkip.Click, AddressOf BtnSkip_Click
        tCell.Controls.Add(BtnSkip)
        If (act = "skip") Then
            BtnSkip.Visible = True
        Else
            BtnSkip.Visible = False
        End If

        LabelLast = New Label
        LabelLast.ID = "label_last"
        LabelLast.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(LabelLast)
        BtnLast = New Button
        BtnLast.ID = "btn_last"
        BtnLast.Width = 60
        BtnLast.Text = "上一筆"
        AddHandler BtnLast.Click, AddressOf BtnLast_Click
        tCell.Controls.Add(BtnLast)

        LabelCount = New Label
        LabelCount.ID = "label_count"
        LabelCount.Font.Bold = True
        If (signoffalready = False) Then '如果點選之簽核通知單為已覆核過, 則不顯示筆數 
            If (actmode = "") Then
                LabelCount.Text = "&nbsp&nbsp第&nbsp&nbsp 1/1 &nbsp&nbsp筆"
            Else
                If (Session("sgcount") <> 0) Then
                    LabelCount.Text = "&nbsp&nbsp第&nbsp&nbsp" & ((Session("startindex") + 1) & "/" & Session("sgcount")) & "&nbsp&nbsp筆"
                Else
                    LabelCount.Text = "&nbsp&nbsp第&nbsp&nbsp" & (Session("startindex") + 1) & "/1&nbsp&nbsp筆"
                End If
            End If
        End If
        tCell.Controls.Add(LabelCount)

        LabelNext = New Label
        LabelNext.ID = "label_next"
        LabelNext.Text = "&nbsp&nbsp"
        tCell.Controls.Add(LabelNext)
        BtnNext = New Button
        BtnNext.ID = "btn_next"
        BtnNext.Width = 60
        BtnNext.Text = "下一筆"
        AddHandler BtnNext.Click, AddressOf BtnNext_Click
        tCell.Controls.Add(BtnNext)

        SignOffStatusLabel = New Label
        SignOffStatusLabel.ID = "label_signoffstatus"
        If (signofftype <> 2) Then '若為未送審,則用XSPWT 一定是無法找到(要用XASCH),所以signofftype會為0(那剛好0是所要的,故就不再另外處理)
            SignOffStatusLabel.Text = info
        Else
            If (sign_status = 1) Then
                SignOffStatusLabel.Text = "因你被設定為此單知悉人,只能觀看"
            Else
                SignOffStatusLabel.Text = "你已知悉過此單"
            End If
        End If
        tCell.Controls.Add(SignOffStatusLabel)

        'LabelPdf = New Label
        'LabelPdf.ID = "label_pdf"
        'LabelPdf.Text = "&nbsp&nbsp"
        'tCell.Controls.Add(LabelPdf)
        'BtnPdf = New Button
        'BtnPdf.ID = "btn_pdf"
        'BtnPdf.Width = 60
        'BtnPdf.Text = "轉Pdf"
        'AddHandler BtnPdf.Click, AddressOf BtnPdf_Click
        'tCell.Controls.Add(BtnPdf)

        BtnSend.Visible = False
        BtnDel.Visible = False
        BtnSave.Visible = False
        BtnCancel.Visible = False
        BtnApproval.Visible = False
        BtnReject.Visible = False
        BtnRecall.Visible = False
        BtnArchieve.Visible = False
        Labelsend.Visible = False
        Labeldel.Visible = False
        Labelapproval.Visible = False
        Labelreject.Visible = False
        Labelrecall.Visible = False
        Labelcancel.Visible = False
        Labelarchieve.Visible = False
        SignOffStatusLabel.Visible = True
        BtnBeInformed.Visible = False
        LabelBeInformed.Visible = False
        LabelFileu.Visible = False
        FileUL.Visible = False
        LabelUpfile.Visible = False
        ChkDel.Visible = False
        BtnFileAct.Visible = False
        ChkFileType.Visible = False
        BtnNext.Visible = False
        BtnSuspend.Visible = False
        BtnLast.Visible = False
        LabelNext.Visible = False
        LabelLast.Visible = False
        LabelCount.Visible = False

        If (docstatus = "D" Or docstatus = "A" Or docstatus = "E" Or ((docstatus = "R" Or docstatus = "B") And formstatusindex = 0 And sign_status = 1)) Then
            LabelFileu.Visible = True
            FileUL.Visible = True
            LabelUpfile.Visible = True
            ChkDel.Visible = True
            BtnFileAct.Visible = True
            ChkFileType.Visible = True
            If (CommSignOff.IsSelfForm(sfid) = 1) Then
                ChkFileType.Checked = True
                ChkFileType.Enabled = False
            Else
                ChkFileType.Checked = False
                ChkFileType.Enabled = True
            End If
        End If

        If (docstatus = "D" Or docstatus = "A" Or docstatus = "E") Then
            If (docstatus = "D") Then
                BtnSend.Visible = True
                Labelsend.Visible = True
                ChkReturn.Visible = False
            End If
            If (docstatus <> "A") Then
                BtnDel.Visible = True
                Labeldel.Visible = True
            End If
            BtnSave.Visible = True
        End If
        If ((docstatus = "R" Or docstatus = "B") And formstatusindex = 0 And sign_status = 1) Then
            BtnSend.Visible = True
            ChkReturn.Visible = False
            Labelsend.Visible = True
            BtnSend.Text = "再送審"
            BtnCancel.Visible = True
            Labelcancel.Visible = True
            BtnSave.Visible = True
        End If

        If (maindocnum = 0) Then
            If (docstatus = "O" And formstatusindex = 0 And sign_status = 1) Then
                BtnApproval.Visible = True
                Labelapproval.Visible = True
                BtnReject.Visible = True
                Labelreject.Visible = True
            End If
            If (docstatus = "F" And sign_status = 1 And signofftype = 1) Then
                BtnArchieve.Visible = True
                Labelarchieve.Visible = True
            End If
            If (signofftype = 2 And sign_status = 1) Then
                BtnBeInformed.Visible = True
                LabelBeInformed.Visible = True
            End If
            Dim nextseq, status As Integer
            If (docstatus <> "A" And docstatus <> "E" And docstatus <> "D" And docstatus <> "F") Then
                SqlCmd = "Select status,seq from [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum & " and uid='" & Session("s_id") & "'"
                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                If (dr.HasRows) Then
                    dr.Read()
                    status = dr(0)
                    nextseq = dr(1) + 1
                End If
                dr.Close()
                connsap.Close()
                If (status = 2 Or status = 100) Then
                    SqlCmd = "Select status from [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum & " and seq=" & nextseq
                    dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                    If (dr.HasRows) Then
                        dr.Read()
                        status = dr(0)
                        dr.Close()
                        connsap.Close()
                        If (status = 1) Then '我已核 , 但下一關還未核 , 故可抽回
                            BtnRecall.Visible = True
                            Labelrecall.Visible = True
                            ChkReturn.Visible = False
                        End If
                    End If
                    dr.Close()
                    connsap.Close()
                End If
            End If
            BtnNext.Visible = True
            BtnSuspend.Visible = True
            BtnLast.Visible = True
            LabelNext.Visible = True
            LabelLast.Visible = True
            LabelCount.Visible = True
            If (sign_status <> 1) Then
                SignOffStatusLabel.Visible = True
            End If
            'If ((actmode = "recycle" Or actmode = "signoff") And sign_status = 1) Then
            Dim mes(2) As String
            If ((actmode = "recycle" Or actmode = "signoff" Or actmode = "recycle_login" Or actmode = "signoff_login")) Then
                If ((Session("startindex") = (Session("sgcount") - 1)) Or Session("sgcount") = 0) Then
                    BtnNext.Enabled = False
                    BtnSuspend.Enabled = False
                    'BtnLast.Enabled = True
                End If
                If (Session("startindex") = 0) Then
                    'BtnNext.Enabled = True
                    BtnLast.Enabled = False
                End If
                If (Session("sgcount") = 0) Then '已簽核過
                    mes = CommSignOff.FormStatusMes(docnum, docstatus)
                    SignOffStatusLabel.Text = mes(1)
                    docstatus = mes(0)
                    signoffalready = True
                End If
            Else
                BtnNext.Enabled = False
                BtnSuspend.Enabled = False
                BtnLast.Enabled = False
                'SignOffStatusLabel.Visible = True
            End If
            If (signoffalready = True) Then '如果點選之簽核通知單為已覆核過, 則不顯示
                BtnNext.Visible = False
                BtnSuspend.Visible = False
                BtnLast.Visible = False
                LabelNext.Visible = False
                LabelLast.Visible = False
                LabelCount.Visible = False
            End If
        End If
        tRow.Cells.Add(tCell)
        FT_m.Rows.Add(tRow)
        ViewState("docstatus") = docstatus
    End Sub
    Sub AddTCreate()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim Labelx As Label
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader

        tRow = New TableRow() 'ronnn
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.Wrap = False
        tCell.HorizontalAlign = HorizontalAlign.Left
        ChkUsingAttach = New CheckBox
        ChkUsingAttach.ID = "chk_usingattach"
        ChkUsingAttach.Text = "使用附檔"
        ChkUsingAttach.AutoPostBack = True
        AddHandler ChkUsingAttach.CheckedChanged, AddressOf ChkUsingAttach_CheckedChanged
        tCell.Controls.Add(ChkUsingAttach)
        Labelx = New Label
        Labelx.ID = "label_s1"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp料號:"
        tCell.Controls.Add(Labelx)

        TxtItemcode = New TextBox
        TxtItemcode.ID = "txt_itemcode"
        TxtItemcode.Width = 150
        TxtItemcode.AutoPostBack = True
        AddHandler TxtItemcode.TextChanged, AddressOf TxtItemcode_TextChanged
        tCell.Controls.Add(TxtItemcode)
        Labelx = New Label
        Labelx.ID = "label_s2"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp說明:"
        tCell.Controls.Add(Labelx)

        TxtItemname = New TextBox
        TxtItemname.ID = "txt_itemname"
        TxtItemname.Width = 200
        tCell.Controls.Add(TxtItemname)
        Labelx = New Label
        Labelx.ID = "label_s3"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp數量:"
        tCell.Controls.Add(Labelx)

        TxtQty = New TextBox
        TxtQty.ID = "txt_qty"
        TxtQty.Width = 30
        tCell.Controls.Add(TxtQty)
        Labelx = New Label
        Labelx.ID = "label_s4"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp金額:"
        tCell.Controls.Add(Labelx)

        TxtUnitPrice = New TextBox
        TxtUnitPrice.ID = "txt_unitprice"
        TxtUnitPrice.Width = 50
        tCell.Controls.Add(TxtUnitPrice)
        Labelx = New Label
        Labelx.ID = "label_s5"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)

        DDLMethod = New DropDownList
        DDLMethod.ID = "ddl_method"
        DDLMethod.Width = 120
        DDLMethod.Items.Clear()
        If (sfid = 51) Then
            DDLMethod.Items.Add("處置方式")
            DDLMethod.Items.Add("生產倉調撥RD倉")
            DDLMethod.Items.Add("RD倉調撥生產倉")
            DDLMethod.Items.Add("領用出庫")
            DDLMethod.Items.Add("一般入庫")
            DDLMethod.Items.Add("其它-請備註說明")
        ElseIf (sfid = 50) Then
            DDLMethod.Items.Add("處置方式")
            DDLMethod.Items.Add("領用 - 出庫備品倉")
            DDLMethod.Items.Add("需求 - 調撥入庫備品倉")
            DDLMethod.Items.Add("入庫後再出庫備品倉")
            DDLMethod.Items.Add("改機 - 入庫備品倉")
            DDLMethod.Items.Add("改機 - 寄回台北_入庫生產倉")
            DDLMethod.Items.Add("返還 - 寄回台北_備品倉調撥入庫生產倉")
            DDLMethod.Items.Add("其它-請備註說明")
        ElseIf (sfid = 49) Then
            DDLMethod.Items.Add("報廢原因")
            DDLMethod.Items.Add("呆料-逾三年未用")
            DDLMethod.Items.Add("設變-已不會再用")
            DDLMethod.Items.Add("損壞-無法維修")
            DDLMethod.Items.Add("損壞-維修費太貴")
            DDLMethod.Items.Add("損壞-已不使用")
            DDLMethod.Items.Add("其它-請備註說明")
        ElseIf (sfid = 100) Then
            DDLMethod.Items.Add("處置方式")
            DDLMethod.Items.Add("取消需求")
            DDLMethod.Items.Add("增加需求")
            DDLMethod.Items.Add("其它-請備註說明")
        ElseIf (sfid = 23) Then
            DDLMethod.Items.Add("離倉原因")
            DDLMethod.Items.Add("借出-廠內")
            DDLMethod.Items.Add("借出-QC檢驗")
            DDLMethod.Items.Add("借出-公司")
            DDLMethod.Items.Add("借出-廠商")
            DDLMethod.Items.Add("送修")
            DDLMethod.Items.Add("品檢驗退")
            DDLMethod.Items.Add("暫放廠商")
            DDLMethod.Items.Add("拆卸待還")
            DDLMethod.Items.Add("其它-請備註說明")
        ElseIf (sfid = 24) Then
            DDLMethod.Items.Add("借入原因")
            DDLMethod.Items.Add("借入-廠商")
            DDLMethod.Items.Add("借入-公司部門")
            DDLMethod.Items.Add("借入-客戶")
            DDLMethod.Items.Add("借入-機台")
            DDLMethod.Items.Add("借入-個人")
            DDLMethod.Items.Add("其它-請備註說明")
        Else
            DDLMethod.Items.Add("處置方式")
        End If
        tCell.Controls.Add(DDLMethod)
        Labelx = New Label
        Labelx.ID = "label_s6"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp備註:"
        tCell.Controls.Add(Labelx)

        TxtNote = New TextBox
        TxtNote.ID = "txt_note"
        TxtNote.Width = 120
        tCell.Controls.Add(TxtNote)
        Labelx = New Label
        Labelx.ID = "label_s7"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)

        BtnAction = New Button
        BtnAction.ID = "btn_action"
        BtnAction.Width = 50
        'BtnAction.Text = "新增"
        If (act = "material_modify") Then
            BtnAction.Text = "修改"
        Else
            BtnAction.Text = "新增"
        End If
        AddHandler BtnAction.Click, AddressOf BtnAction_Click
        tCell.Controls.Add(BtnAction)
        Labelx = New Label
        Labelx.ID = "label_s8"
        Labelx.Text = "&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        BtnReset = New Button
        BtnReset.ID = "btn_reset"
        BtnReset.Width = 50
        BtnReset.Text = "重置"
        AddHandler BtnReset.Click, AddressOf BtnReset_Click
        tCell.Controls.Add(BtnReset)

        tRow.Cells.Add(tCell)
        AddT.Rows.Add(tRow)
        If (act = "material_modify") Then
            SqlCmd = "Select itemcode,itemname,quantity,price,method,comment FROM [dbo].[@XSMLS] T0 WHERE T0.[num] =" & Request.QueryString("num")
            drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
            If (drL.HasRows) Then
                drL.Read()
                TxtItemcode.Text = drL(0)
                TxtItemname.Text = drL(1)
                TxtQty.Text = drL(2)
                TxtUnitPrice.Text = drL(3)
                DDLMethod.SelectedValue = drL(4)
                TxtNote.Text = drL(5)
            End If
            drL.Close()
            connL.Close()
        End If
        '
        tRow = New TableRow()
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.Wrap = False
        tCell.HorizontalAlign = HorizontalAlign.Left

        Labelx = New Label()
        Labelx.ID = "label_selectfile"
        Labelx.Text = "此處為以Excel檔輸入料件:&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp選擇料件檔案"
        tCell.Controls.Add(Labelx)
        Dim FileSel As New FileUpload()
        FileSel.ID = "filesel"
        If (docstatus <> "A") Then
            FileSel.Enabled = True
        Else
            FileSel.Enabled = False
        End If
        tCell.Controls.Add(FileSel)

        Labelx = New Label
        Labelx.ID = "label_s51"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)

        DDLMethod1 = New DropDownList
        DDLMethod1.ID = "ddl_method1"
        DDLMethod1.Width = 140
        DDLMethod1.Items.Clear()
        If (sfid = 51) Then
            DDLMethod1.Items.Add("處置方式")
            DDLMethod1.Items.Add("生產倉調撥RD倉")
            DDLMethod1.Items.Add("RD倉調撥生產倉")
            DDLMethod1.Items.Add("領用出庫")
            DDLMethod1.Items.Add("一般入庫")
        ElseIf (sfid = 50) Then
            DDLMethod1.Items.Add("處置方式")
            DDLMethod1.Items.Add("領用 - 出庫備品倉")
            DDLMethod1.Items.Add("需求 - 調撥入庫備品倉")
            DDLMethod1.Items.Add("入庫後再出庫備品倉")
            DDLMethod1.Items.Add("改機 - 入庫備品倉")
            DDLMethod1.Items.Add("改機 - 寄回台北_入庫生產倉")
            DDLMethod1.Items.Add("返還 - 寄回台北_備品倉調撥入庫生產倉")
        ElseIf (sfid = 49) Then
            DDLMethod1.Items.Add("報廢原因")
            DDLMethod1.Items.Add("呆料-逾三年未用")
            DDLMethod1.Items.Add("設變-已不會再用")
            DDLMethod1.Items.Add("損壞-無法維修")
            DDLMethod1.Items.Add("損壞-維修費太貴")
            DDLMethod1.Items.Add("損壞-已不使用")
            DDLMethod1.Items.Add("其它-備註說明")
        ElseIf (sfid = 100) Then
            DDLMethod1.Items.Add("處置方式")
            DDLMethod1.Items.Add("取消需求")
            DDLMethod1.Items.Add("增加需求")
        ElseIf (sfid = 23) Then
            DDLMethod1.Items.Add("離倉原因")
            DDLMethod1.Items.Add("借出-廠內")
            DDLMethod1.Items.Add("借出-公司")
            DDLMethod1.Items.Add("借出-廠商")
            DDLMethod1.Items.Add("送修")
            DDLMethod1.Items.Add("品檢驗退")
            DDLMethod1.Items.Add("暫放廠商")
            DDLMethod1.Items.Add("其它-備註說明")
        ElseIf (sfid = 24) Then
            DDLMethod1.Items.Add("借入原因")
            DDLMethod1.Items.Add("借入-廠商")
            DDLMethod1.Items.Add("借入-公司部門")
            DDLMethod1.Items.Add("借入-客戶")
            DDLMethod1.Items.Add("借入-機台")
            DDLMethod1.Items.Add("借入-個人")
            DDLMethod1.Items.Add("其它-請備註說明")
        Else
            DDLMethod1.Items.Add("處置方式")
        End If
        tCell.Controls.Add(DDLMethod1)

        Labelx = New Label()
        Labelx.ID = "label_selectfile1"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        Dim BtnImportMaterial As New Button
        BtnImportMaterial.ID = "btn_ImportMaterial"
        BtnImportMaterial.Text = "導入料件"
        AddHandler BtnImportMaterial.Click, AddressOf BtnImportMaterial_Click
        tCell.Controls.Add(BtnImportMaterial)

        tRow.Cells.Add(tCell)

        AddT.Rows.Add(tRow)
    End Sub
    'Protected Sub BtnImportMaterial_Click(ByVal sender As Object, ByVal e As EventArgs) ' Using Microsoft Excel API
    '    Dim oExcel As Excel.Application
    '    Dim oBook As Excel.Workbook
    '    Dim oBooks As Excel.Workbooks
    '    Dim oSheet As Excel.Worksheet
    '    Dim filepath, str() As String
    '    Dim exitflag As Boolean
    '    exitflag = False
    '    'Dim FileSel As FileUpload()
    '    Dim FileSel As FileUpload = CType(AddT.FindControl("filesel"), FileUpload)
    '    Dim i, count As Integer
    '    i = 2
    '    count = 0
    '    Try
    '        If (FileSel.HasFile) Then
    '            If (DDLMethod1.SelectedIndex = 0) Then
    '                CommUtil.ShowMsg(Me, "原因要選擇")
    '                Exit Sub
    '            End If
    '            filepath = Application("localdir") & "FileTemp\" & FileSel.FileName
    '            str = Split(FileSel.FileName, ".")
    '            If (str(1) <> "xls" And str(1) <> "xlsx") Then
    '                CommUtil.ShowMsg(Me, "要上傳檔案需為xls or xlsx")
    '                Exit Sub
    '            End If
    '            FileSel.SaveAs(filepath)
    '            '建立Excel物件並開啟C:\01.xls中的Sheet1
    '            oExcel = CreateObject("Excel.Application")
    '            oExcel.Visible = True
    '            oBooks = oExcel.Workbooks
    '            oBook = oBooks.Open(filepath)
    '            Try
    '                oSheet = oBook.Worksheets("工作表1")
    '                If (IsNumeric(oSheet.Cells(1, 3).value)) Then
    '                    CommUtil.ShowMsg(Me, "料件應從Excel第二列開始,請重新編輯")
    '                    exitflag = True
    '                End If
    '                If (exitflag = False) Then
    '                    Do While oSheet.Cells(i, 1).value <> ""
    '                        If (Not IsNumeric(oSheet.Cells(i, 3).value)) Then
    '                            CommUtil.ShowMsg(Me, "Excel 第" & i & "列之數量不是整數")
    '                            exitflag = True
    '                            Exit Do
    '                        End If
    '                        SqlCmd = "select T0.itemname from dbo.OITM T0 where T0.itemcode='" & oSheet.Cells(i, 1).value & "'"
    '                        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
    '                        If (dr.HasRows) Then

    '                        Else
    '                            CommUtil.ShowMsg(Me, "Sap無" & oSheet.Cells(i, 1).value & "這料件")
    '                            dr.Close()
    '                            connsap.Close()
    '                            exitflag = True
    '                            Exit Do
    '                        End If
    '                        dr.Close()
    '                        connsap.Close()
    '                        count = count + 1
    '                        i = i + 1
    '                    Loop
    '                End If
    '                If (exitflag = False) Then
    '                    i = 2
    '                    Dim mprice As Double
    '                    Dim mitemname, method, itemcode As String
    '                    Dim qty As Integer
    '                    method = DDLMethod1.SelectedValue
    '                    
    '                    Do While oSheet.Cells(i, 1).value <> ""
    '                        mitemname = oSheet.Cells(i, 2).value
    '                        itemcode = oSheet.Cells(i, 1).value
    '                        qty = oSheet.Cells(i, 3).value
    '                        SqlCmd = "select T0.itemname from dbo.OITM T0 where T0.itemcode='" & itemcode & "'"
    '                        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
    '                        If (dr.HasRows) Then
    '                            dr.Read()
    '                            mitemname = dr(0)
    '                            mprice = GetUnitPrice(itemcode)
    '                        End If
    '                        dr.Close()
    '                        connsap.Close()
    '                        SqlCmd = "insert into [dbo].[@XSMLS] (docentry,itemcode,itemname,quantity,method,price) " &
    '                    "values(" & docnum & ",'" & itemcode & "','" & mitemname & "'," & qty & ",'" & method & "'," & mprice & ")"
    '                        CommUtil.SqlSapExecute("ins", SqlCmd, connsap)
    '                        connsap.Close()
    '                        i = i + 1
    '                    Loop
    '                    If (count = 0) Then
    '                        CommUtil.ShowMsg(Me, "Excel 內無任何導入之料件,請Check")
    '                    End If
    '                End If
    '            Catch ex As Exception
    '                CommUtil.ShowMsg(Me, "無工作表1名稱之工作表")
    '            End Try

    '            '關閉並釋放Excel物件
    '            oBook.Close(False)
    '            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook)
    '            oBook = Nothing
    '            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBooks)
    '            oBooks = Nothing
    '            oExcel.Quit()
    '            System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel)
    '            oExcel = Nothing
    '            IO.File.Delete(filepath)
    '            Response.Redirect("cLsignoff.aspx?smid=sg&smode=2&status=" & docstatus &
    '            "&docnum=" & docnum & "&formstatusindex=" & formstatusindex &
    '            "&formtypeindex=" & formtypeindex & "&sfid=" & sfid & "&subject=" & TxtSubject.Text & "&signflowmode=" & signflowmode)
    '        Else
    '            CommUtil.ShowMsg(Me, "未指定上傳檔案")
    '        End If

    '    Catch ex As Exception
    '        CommUtil.ShowMsg(Me, "操作Excel遇到問題")
    '    End Try
    'End Sub

    Protected Sub BtnImportMaterial_Click(ByVal sender As Object, ByVal e As EventArgs) 'Using Open Source API
        Dim filepath, str() As String
        Dim exitflag As Boolean
        Dim readStream As IO.FileStream = Nothing
        Dim ds As New DataSet
        Dim firstchara As String
        exitflag = False
        'Dim FileSel As FileUpload()
        Dim FileSel As FileUpload = CType(AddT.FindControl("filesel"), FileUpload)
        Dim i, count As Integer
        i = 1
        count = 0
        Try
            If (FileSel.HasFile) Then
                If (DDLMethod1.SelectedIndex = 0) Then
                    CommUtil.ShowMsg(Me, "原因要選擇")
                    Exit Sub
                End If
                filepath = Application("localdir") & "FileTemp\" & FileSel.FileName
                str = Split(FileSel.FileName, ".")
                If (str(1) <> "xls" And str(1) <> "xlsx") Then
                    CommUtil.ShowMsg(Me, "要上傳檔案需為xls or xlsx")
                    Exit Sub
                End If
                FileSel.SaveAs(filepath)
                '開啟Excel之filestream , 並建立excel reader , 讀取的是第一個sheet
                Dim reader As IExcelDataReader
                readStream = New IO.FileStream(filepath, IO.FileMode.Open)
                If (str(1) = "xls") Then
                    reader = ExcelReaderFactory.CreateBinaryReader(readStream, New ExcelReaderConfiguration())
                ElseIf (str(1) = "xlsx") Then
                    reader = ExcelReaderFactory.CreateOpenXmlReader(readStream, New ExcelReaderConfiguration())
                End If
                ds = reader.AsDataSet(New ExcelDataSetConfiguration())
                firstchara = Left(ds.Tables(0).Rows(0)(0), 1)
                If (IsNumeric(firstchara)) Then
                    CommUtil.ShowMsg(Me, "料件應從Excel第二列開始,請重新編輯")
                    exitflag = True
                End If
                If (exitflag = False) Then
                    For i = 1 To ds.Tables(0).Rows.Count - 1
                        If (Not IsNumeric(ds.Tables(0).Rows(i)(2))) Then
                            CommUtil.ShowMsg(Me, "Excel 第" & i + 1 & "列之數量不是整數")
                            exitflag = True
                            Exit For
                        End If
                        SqlCmd = "select T0.itemname from dbo.OITM T0 where T0.itemcode='" & ds.Tables(0).Rows(i)(0) & "'"
                        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                        If (dr.HasRows) Then

                        Else
                            CommUtil.ShowMsg(Me, "Sap無" & ds.Tables(0).Rows(i)(0) & "這料件")
                            dr.Close()
                            connsap.Close()
                            exitflag = True
                            Exit For
                        End If
                        dr.Close()
                        connsap.Close()
                        count = count + 1
                    Next
                End If
                If (exitflag = False) Then
                    i = 1
                    Dim mprice As Double
                    Dim mitemname, method, itemcode As String
                    Dim qty As Integer
                    method = DDLMethod1.SelectedValue
                    For i = 1 To ds.Tables(0).Rows.Count - 1
                        mitemname = ds.Tables(0).Rows(i)(1)
                        itemcode = ds.Tables(0).Rows(i)(0)
                        qty = ds.Tables(0).Rows(i)(2)
                        SqlCmd = "select T0.itemname from dbo.OITM T0 where T0.itemcode='" & itemcode & "'"
                        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                        If (dr.HasRows) Then
                            dr.Read()
                            mitemname = dr(0)
                            mprice = GetUnitPrice(itemcode)
                        End If
                        dr.Close()
                        connsap.Close()
                        SqlCmd = "insert into [dbo].[@XSMLS] (docentry,itemcode,itemname,quantity,method,price) " &
                    "values(" & docnum & ",'" & itemcode & "','" & mitemname & "'," & qty & ",'" & method & "'," & mprice & ")"
                        CommUtil.SqlSapExecute("ins", SqlCmd, connsap)
                        connsap.Close()
                    Next
                End If
                '關閉並釋放Excel物件
                readStream.Close()
                reader.Close()
                ds.Dispose()
                IO.File.Delete(filepath)
                If (count = 0) Then
                    CommUtil.ShowMsg(Me, "Excel 內無任何導入之料件,請Check")
                    Exit Sub
                End If
                Response.Redirect("cLsignoff.aspx?smid=sg&smode=2&status=" & docstatus &
                            "&docnum=" & docnum & "&formstatusindex=" & formstatusindex &
                            "&formtypeindex=" & formtypeindex & "&sfid=" & sfid & "&subject=" & TxtSubject.Text & "&signflowmode=" & signflowmode)
            Else
                CommUtil.ShowMsg(Me, "未指定上傳檔案")
            End If

        Catch ex As Exception
            CommUtil.ShowMsg(Me, "操作Excel遇到問題")
        End Try
    End Sub
    Function TxtReasonChangeOfSfid49_50_51()
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        TxtReasonChangeOfSfid49_50_51 = False
        SqlCmd = "Select descrip FROM [dbo].[@XSMLS] T0 WHERE T0.[docentry] =" & docnum & " and head=1"
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            drL.Read()
            If (drL(0) <> TxtReason.Text) Then
                TxtReasonChangeOfSfid49_50_51 = True
            End If
        End If
        drL.Close()
        connL.Close()
    End Function
    Protected Sub BtnAction_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim num As Long
        Dim itemcode, itemname, method, comment As String
        Dim qty As Integer
        Dim price As Double
        If (TxtReasonChangeOfSfid49_50_51()) Then
            CommUtil.ShowMsg(Me, "事由說明欄已變更,請先儲存再新增/修改")
            Exit Sub
        End If
        itemcode = TxtItemcode.Text
        itemname = TxtItemname.Text
        method = DDLMethod.SelectedValue
        comment = TxtNote.Text
        If (Not IsNumeric(TxtQty.Text)) Then
            CommUtil.ShowMsg(Me, "數量欄位必須是數字")
            Exit Sub
        End If
        If (Not IsNumeric(TxtUnitPrice.Text)) Then
            CommUtil.ShowMsg(Me, "價格欄位必須是數字")
            Exit Sub
        End If
        If (itemcode = "" Or itemname = "") Then
            CommUtil.ShowMsg(Me, "料號及說明欄不能空白")
            Exit Sub
        End If
        If (DDLMethod.SelectedIndex = 0) Then
            If (sfid = 51 Or sfid = 50 Or sfid = 100) Then
                CommUtil.ShowMsg(Me, "處置方式必須選擇")
            ElseIf (sfid = 49) Then
                CommUtil.ShowMsg(Me, "報廢原因必須選擇")
            ElseIf (sfid = 23) Then
                CommUtil.ShowMsg(Me, "離倉原因必須選擇")
            ElseIf (sfid = 24) Then
                CommUtil.ShowMsg(Me, "借入原因必須選擇")
            End If
            Exit Sub
        Else
            If (DDLMethod.SelectedValue = "其它-請備註說明") Then
                If (comment = "") Then
                    CommUtil.ShowMsg(Me, "你選擇了 '其它-請備註說明', 故備註欄需填寫原因")
                    Exit Sub
                End If
            End If
        End If
        Dim mtype As Integer
        mtype = 0
        SqlCmd = "Select mtype FROM [dbo].[@XSMLS] T0 WHERE T0.[docentry] =" & docnum & " and head=1"
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            drL.Read()
            mtype = drL(0)
        End If
        drL.Close()
        connL.Close()
        'If (mtype = 1) Then
        '    If (comment = "") Then
        '        CommUtil.ShowMsg(Me, "請在備註欄註明-料件是何人帶走到何處")
        '        Exit Sub
        '    End If
        'End If
        qty = CInt(TxtQty.Text)
        price = CDbl(TxtUnitPrice.Text)
        If (sender.Text = "新增") Then
            SqlCmd = "insert into [dbo].[@XSMLS] (docentry,itemcode,itemname,quantity,method,price,comment,mtype) " &
        "values(" & docnum & ",'" & itemcode & "','" & itemname & "'," & qty & ",'" & method & "'," & price & ",'" & comment & "'," & mtype & ")"
            CommUtil.SqlSapExecute("ins", SqlCmd, connsap)
            connsap.Close()
            If (TxtSubject.Text = "") Then
                If (sfid = 23) Then
                    TxtSubject.Text = method & "-料件出倉事宜"
                ElseIf (sfid = 24) Then
                    TxtSubject.Text = method & "-料件借入事宜"
                End If
            End If
        ElseIf (sender.Text = "修改") Then
            num = Request.QueryString("num")
            SqlCmd = "update [dbo].[@XSMLS] set itemcode='" & itemcode & "',itemname='" & itemname & "'," &
                     "quantity=" & qty & ",method='" & method & "',price=" & price & ",comment='" & comment & "' " &
                    "where num=" & num
            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
            connsap.Close()
            sender.Text = "新增"
        End If
        If (docstatus = "D") Then
            docstatus = "E" '可能動到price , 故改狀態(就是要user再按一次儲存 , 已確保此price 與料件總價一致)
            BtnSend.Visible = False
            SqlCmd = "update [dbo].[@XASCH] set status='E' " &
            " where docnum=" & docnum
            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
            connsap.Close()
        End If
        Response.Redirect("cLsignoff.aspx?smid=sg&smode=2&status=" & docstatus &
        "&docnum=" & docnum & "&formstatusindex=" & formstatusindex &
        "&formtypeindex=" & formtypeindex & "&sfid=" & sfid & "&subject=" & TxtSubject.Text & "&signflowmode=" & signflowmode)
    End Sub
    Protected Sub BtnReset_Click(ByVal sender As Object, ByVal e As EventArgs)
        If (TxtReasonChangeOfSfid49_50_51()) Then
            CommUtil.ShowMsg(Me, "事由說明欄已變更,請先儲存再重置")
            Exit Sub
        End If
        If (sfid <> 23) Then
            ShowMaterialList(0)
        Else
            ShowMaterialList23_24(0)
        End If
        AddTFieldClear()
    End Sub
    Sub AddTFieldClear()
        TxtItemcode.Text = ""
        TxtItemname.Text = ""
        DDLMethod.SelectedIndex = 0
        TxtQty.Text = ""
        TxtUnitPrice.Text = ""
        TxtNote.Text = ""
        BtnAction.Text = "新增"
    End Sub
    Sub TxtItemcode_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
        If (TxtItemcode.Text <> "") Then
            TxtNote.Text = ""
            TxtUnitPrice.Text = ""
            If (TxtItemcode.Text <> "All" And TxtItemcode.Text <> "ALL") Then
                SqlCmd = "select T0.itemname from dbo.OITM T0 where T0.itemcode='" & TxtItemcode.Text & "'"
                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                If (dr.HasRows) Then
                    dr.Read()
                    TxtItemname.Text = dr(0)
                    TxtUnitPrice.Text = GetUnitPrice(TxtItemcode.Text)
                Else
                    CommUtil.ShowMsg(Me, "Sap無" & TxtItemcode.Text & "這料件")
                    'Exit Sub
                End If
                dr.Close()
                connsap.Close()
            Else
                TxtQty.Text = 1
                TxtItemname.Text = "一批料件,詳附檔"
                TxtNote.Text = "料件太多,以附檔呈現"
            End If
        End If
    End Sub
    Function GetUnitPrice(itemcode As String)
        Dim price As Double
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        SqlCmd = "Select IsNull(T0.[Price],0) FROM dbo.POR1 T0 WHERE T0.[ItemCode] ='" & itemcode & "' ORDER BY docentry desc"
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        drL.Read()
        If (drL.HasRows) Then
            If (drL(0) <> 0) Then
                price = drL(0)
            Else
                drL.Close()
                connL.Close()
                '廠商價格
                SqlCmd = "SELECT IsNull(T0.[Price],0) " &
                "FROM ITM1 T0 WHERE T0.[ItemCode] ='" & TxtItemcode.Text & "' and T0.[PriceList] =7"
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                drL.Read()
                If (drL(0) <> 0) Then
                    price = drL(0)
                Else
                    drL.Close()
                    connL.Close()
                    '標準價格
                    SqlCmd = "SELECT IsNull(T0.[Price],0) " &
                    "FROM ITM1 T0 WHERE T0.[ItemCode] ='" & TxtItemcode.Text & "' and T0.[PriceList] =1"
                    drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                    drL.Read()
                    If (drL(0) <> 0) Then
                        price = drL(0)
                    Else
                        drL.Close()
                        connL.Close()
                        '平均價 from C02
                        SqlCmd = "SELECT T0.[AvgPrice] " &
                        "FROM OITW T0 WHERE T0.[ItemCode] ='" & TxtItemcode.Text & "' and T0.[WhsCode] = 'C02'"
                        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                        drL.Read()
                        If (drL(0) <> 0) Then
                            price = drL(0)
                        Else
                            drL.Close()
                            connL.Close()
                            '平均價 from C01
                            SqlCmd = "SELECT T0.[AvgPrice] " &
                            "FROM OITW T0 WHERE T0.[ItemCode] ='" & TxtItemcode.Text & "' and T0.[WhsCode] = 'C01'"
                            drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                            drL.Read()
                            If (drL(0) <> 0) Then
                                price = drL(0)
                            Else
                                price = 0
                                CommUtil.ShowMsg(Me, "於sap系統查不到料件:" & TxtItemcode.Text & "之金額")
                            End If
                        End If
                    End If
                End If
            End If
        End If
        drL.Close()
        connL.Close()
        Return price
    End Function
    Sub InitTableToSfid12()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim tImage As Image
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim tTxt As TextBox
        Dim Labelx As Label
        Dim dDDL As DropDownList

        Dim ce As CalendarExtender
        Dim rRBL As RadioButtonList
        Dim cChk As CheckBox
        Dim BColor As Drawing.Color
        BColor = System.Drawing.Color.LightBlue
        ContentT.Font.Name = "標楷體"
        ContentT.Font.Size = 14
        tRow = New TableRow()
        'tRow.Font.Bold = True
        For j = 1 To 6
            tCell = New TableCell
            tCell.BorderWidth = 0
            tCell.Width = 200
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Controls.Add(tCell)
        Next
        ContentT.Rows.Add(tRow)
        'row=0 Title
        tRow = New TableRow()
        tRow.Font.Bold = True
        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Left
        tImage = New Image
        tImage.ID = "image_logo"
        tImage.ImageUrl = "~/image/jetlog80%.jpg"
        tCell.Controls.Add(tImage)
        tRow.Controls.Add(tCell)
        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.Font.Size = 36
        tCell.ColumnSpan = 4
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "捷智科技 AOI/SPI 內部發包單"
        tRow.Controls.Add(tCell)
        tCell = New TableCell
        tCell.Font.Size = 12
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Right
        tCell.VerticalAlign = VerticalAlign.Bottom
        If (docnum <> 0) Then
            If (docstatus = "E" Or docstatus = "D") Then
                SqlCmd = "Select convert(varchar(12), docdate, 111) from [dbo].[@XASCH] where docnum=" & docnum
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    tCell.Text = "單號:" & docnum & "<br>建單日期:" & drL(0)
                End If
                drL.Close()
                connL.Close()
            Else
                SqlCmd = "Select convert(varchar(12), signdate, 111) from [dbo].[@XSPWT] where docentry=" & docnum & " and seq=1"
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    tCell.Text = "單號:" & docnum & "<br>送審日期:" & drL(0)
                End If
                drL.Close()
                connL.Close()
            End If
        End If
        tRow.Controls.Add(tCell)
        ContentT.Rows.Add(tRow)
        'Row = 1 cell 1-3
        tRow = New TableRow()
        'tRow.Font.Bold = True
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "發包單位"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'Row = 1 cell 2-3
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Center
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_area"
        rRBL.Items.Add("台北捷智")
        rRBL.Items.Add("深圳捷智通")
        rRBL.Items.Add("昆山捷豐")
        rRBL.RepeatDirection = RepeatDirection.Vertical
        rRBL.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(rRBL)
        tRow.Controls.Add(tCell)
        'row=1 cell 4
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "負責業務"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=1 cell 5-6
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Center
        tTxt = New TextBox()
        tTxt.ID = "txt_sales"
        tTxt.Font.Name = "標楷體"
        tTxt.Font.Size = 14
        tTxt.Width = 100
        tTxt.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(tTxt)
        tRow.Controls.Add(tCell)

        ContentT.Rows.Add(tRow)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'row=2 cell=1
        tRow = New TableRow()
        'tRow.Font.Bold = True
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "捷智機型"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=2 cell 2-3
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Center
        If (Session("usingwhs") = "C02") Then
            SqlCmd = "SELECT T0.u_model,T0.u_mdesc,T0.u_mtype " &
                 "FROM dbo.[@UMMD] T0 where T0.u_model<>'6800' and T0.u_model<>'6500S' and (T0.u_mtype='SPI' or T0.u_mtype='AOI' or " &
                 "T0.u_mtype='3DAOI') order by T0.u_model,T0.u_mcode"
        ElseIf (Session("usingwhs") = "C01") Then
            SqlCmd = "SELECT T0.u_model,T0.u_mdesc,T0.u_mtype " &
                 "FROM dbo.[@UMMD] T0 where T0.u_mtype='ICT' order by T0.u_model,T0.u_mcode"
        Else
            CommUtil.ShowMsg(Me, "倉別設定須為C01 or C02已決定是ICT or AOI")
        End If
        dDDL = New DropDownList()
        dDDL.Items.Add("請選擇")
        If (Session("usingwhs") = "C02" Or Session("usingwhs") = "C01") Then
            drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
            If (drL.HasRows) Then
                Do While (drL.Read())
                    dDDL.Items.Add(drL(0) & "-" & drL(1))
                Loop
            End If
            drL.Close()
            connL.Close()
        End If
        dDDL.ID = "ddl_model"
        dDDL.Font.Name = "標楷體"
        dDDL.Font.Size = 14
        dDDL.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(dDDL)
        tRow.Controls.Add(tCell)
        'row=2 cell=4
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "客戶名稱"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=2 cell=5
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        tTxt = New TextBox()
        tTxt.ID = "txt_customer"
        tTxt.Font.Name = "標楷體"
        tTxt.Font.Size = 14
        tTxt.Width = 130
        tTxt.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(tTxt)
        tRow.Controls.Add(tCell)
        'row=2 cell=6
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        Labelx = New Label
        Labelx.ID = "label_2_6"
        Labelx.Text = "數量:"
        tCell.Controls.Add(Labelx)
        tTxt = New TextBox()
        tTxt.ID = "txt_amount"
        tTxt.Font.Name = "標楷體"
        tTxt.Font.Size = 14
        tTxt.Width = 30
        tTxt.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(tTxt)
        tRow.Controls.Add(tCell)

        ContentT.Rows.Add(tRow)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'row=3 cell=1
        tRow = New TableRow()
        'tRow.Font.Bold = True

        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "出貨型號"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=3 cell=2,3
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Center
        tTxt = New TextBox()
        tTxt.ID = "txt_shipmodel"
        tTxt.Font.Name = "標楷體"
        tTxt.Font.Size = 14
        tTxt.Width = 150
        tTxt.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(tTxt)
        tRow.Controls.Add(tCell)
        'row=3 cell=4
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "工廠出貨日期"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=3 cell=5-6
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Center
        tTxt = New TextBox()
        tTxt.ID = "txt_shipdate"
        tTxt.Font.Name = "標楷體"
        tTxt.Font.Size = 14
        tTxt.Width = 120
        tTxt.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(tTxt)
        ce = New CalendarExtender
        ce.TargetControlID = tTxt.ID
        ce.ID = "ce_shipdate"
        ce.Format = "yyyy/MM/dd"
        tCell.Controls.Add(ce)
        tRow.Controls.Add(tCell)

        ContentT.Rows.Add(tRow)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'row=4 cell=1-6
        tRow = New TableRow()
        'tRow.Font.Bold = True

        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.ColumnSpan = 6
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "系統規格要求"
        tCell.Font.Bold = True
        tRow.Controls.Add(tCell)
        ContentT.Rows.Add(tRow)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'row=5 cell=1
        tRow = New TableRow()
        'tRow.Font.Bold = True

        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "產品大小/重量/厚度"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=5 cell=2,3
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Left
        Labelx = New Label
        Labelx.ID = "label_5_2_1"
        Labelx.Text = "客戶端產品大小:長"
        tCell.Controls.Add(Labelx)
        tTxt = New TextBox()
        tTxt.ID = "txt_uutlength"
        tTxt.Font.Name = "標楷體"
        tTxt.Font.Size = 14
        tTxt.Width = 50
        tTxt.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_5_2_2"
        Labelx.Text = "mm/寬"
        tCell.Controls.Add(Labelx)
        tTxt = New TextBox()
        tTxt.ID = "txt_uutwidth"
        tTxt.Font.Name = "標楷體"
        tTxt.Font.Size = 14
        tTxt.Width = 50
        tTxt.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_5_2_3"
        Labelx.Text = "mm<br>產品重量:"
        tCell.Controls.Add(Labelx)
        tTxt = New TextBox()
        tTxt.ID = "txt_uutweight"
        tTxt.Font.Name = "標楷體"
        tTxt.Font.Size = 14
        tTxt.Width = 50
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_5_2_4"
        Labelx.Text = "KG<br>產品厚度:"
        tCell.Controls.Add(Labelx)
        tTxt = New TextBox()
        tTxt.ID = "txt_uutthick"
        tTxt.Font.Name = "標楷體"
        tTxt.Font.Size = 14
        tTxt.Width = 60
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_5_2_5"
        Labelx.Text = "mm(與皮帶型式有關)"
        tCell.Controls.Add(Labelx)
        tRow.Controls.Add(tCell)
        'row=5 cell 4
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "生產線高度"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=5 cell=5,6
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Left
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_plheight"
        rRBL.Items.Add("900+-20 mm")
        rRBL.Items.Add("其它特殊高度")
        rRBL.RepeatDirection = RepeatDirection.Horizontal
        rRBL.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(rRBL)
        tTxt = New TextBox()
        tTxt.ID = "txt_plotherheight"
        tTxt.Font.Name = "標楷體"
        tTxt.Font.Size = 16
        tTxt.Width = 50
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_5_2_6"
        Labelx.Text = "&nbsp&nbsp+-&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        tTxt = New TextBox()
        tTxt.ID = "txt_plotherheighttol"
        tTxt.Font.Name = "標楷體"
        tTxt.Font.Size = 14
        tTxt.Width = 50
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_5_2_7"
        Labelx.Text = "&nbsp&nbspmm"
        tCell.Controls.Add(Labelx)
        tRow.Controls.Add(tCell)

        ContentT.Rows.Add(tRow)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        tRow = New TableRow()
        ' tRow.Font.Bold = True
        'row=6 cell=1
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "是否有載具"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=6 cell=2,3
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Left
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_withfixture"
        rRBL.Items.Add("無載具")
        rRBL.Items.Add("有載具")
        rRBL.RepeatDirection = RepeatDirection.Horizontal
        rRBL.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(rRBL)
        Labelx = New Label
        Labelx.ID = "label_6_2_1"
        Labelx.Text = "載具大小:"
        tCell.Controls.Add(Labelx)
        tTxt = New TextBox()
        tTxt.ID = "txt_fixturesize"
        tTxt.Font.Name = "標楷體"
        tTxt.Font.Size = 14
        tTxt.Width = 80
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_6_2_2"
        Labelx.Text = "&nbsp&nbspmm"
        tCell.Controls.Add(Labelx)
        tRow.Controls.Add(tCell)
        'row=6 cell=4
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "載具樣式"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=6 cell=5,6
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Left
        tCell.Text = "請詳細確認載具樣式,尤其是與皮帶接觸邊之幾何構形;必要時拍圖片;或向客戶取得圖檔;或寄回實板參考"
        tRow.Controls.Add(tCell)

        ContentT.Rows.Add(tRow)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        tRow = New TableRow()
        'tRow.Font.Bold = True
        'row=7 cell=1
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "進板方向"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=7 cell=2,3
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Left
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_pcbdir"
        rRBL.Items.Add("由左 --> 右")
        rRBL.Items.Add("由右 --> 左")
        rRBL.Items.Add("雙向")
        rRBL.RepeatDirection = RepeatDirection.Vertical
        rRBL.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(rRBL)
        tRow.Controls.Add(tCell)
        'row=7 cell=4
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "Cycle Time"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=7 cell=5,6
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Left
        Labelx = New Label
        Labelx.ID = "label_7_5_1"
        Labelx.Text = "客戶具體板子大小:"
        tCell.Controls.Add(Labelx)
        tTxt = New TextBox()
        tTxt.ID = "txt_pcbsizeX"
        tTxt.Font.Name = "標楷體"
        tTxt.Font.Size = 14
        tTxt.Width = 50
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_7_5_2"
        Labelx.Text = "&nbsp&nbsp*&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        tTxt = New TextBox()
        tTxt.ID = "txt_pcbsizeY"
        tTxt.Font.Name = "標楷體"
        tTxt.Font.Size = 14
        tTxt.Width = 50
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_7_5_3"
        Labelx.Text = "&nbsp&nbspmm<br>"
        tCell.Controls.Add(Labelx)
        tTxt = New TextBox()
        tTxt.ID = "txt_cycletime"
        tTxt.Font.Name = "標楷體"
        tTxt.Font.Size = 14
        tTxt.Width = 50
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_7_5_4"
        Labelx.Text = "&nbsp&nbsp秒<br>"
        tCell.Controls.Add(Labelx)
        tRow.Controls.Add(tCell)

        ContentT.Rows.Add(tRow)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        tRow = New TableRow()
        'tRow.Font.Bold = True
        'row=8 cell=1
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "作業系統"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=8 cell=2,3
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Left
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_oslang"
        rRBL.Items.Add("繁體版")
        rRBL.Items.Add("簡體版")
        rRBL.Items.Add("英文版")
        rRBL.RepeatDirection = RepeatDirection.Vertical
        rRBL.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(rRBL)
        tRow.Controls.Add(tCell)
        'row=8 cell=4
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "加裝Z軸"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=8 cell=5,6
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Left
        cChk = New CheckBox
        cChk.ID = "chk_upz"
        cChk.Text = ""
        tCell.Controls.Add(cChk)
        Labelx = New Label
        Labelx.ID = "label_8_5_1"
        Labelx.Text = "上Z軸行程:"
        tCell.Controls.Add(Labelx)
        tTxt = New TextBox()
        tTxt.ID = "txt_zmm"
        tTxt.Font.Name = "標楷體"
        tTxt.Font.Size = 14
        tTxt.Width = 50
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_8_5_2"
        Labelx.Text = "mm&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)

        cChk = New CheckBox
        cChk.ID = "chk_downz"
        cChk.Text = ""
        tCell.Controls.Add(cChk)
        Labelx = New Label
        Labelx.ID = "label_8_5_3"
        Labelx.Text = "下Z軸行程:"
        tCell.Controls.Add(Labelx)
        tTxt = New TextBox()
        tTxt.ID = "txt_dzmm"
        tTxt.Font.Name = "標楷體"
        tTxt.Font.Size = 14
        tTxt.Width = 50
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_8_5_4"
        Labelx.Text = "mm"
        tCell.Controls.Add(Labelx)
        tRow.Controls.Add(tCell)

        ContentT.Rows.Add(tRow)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        tRow = New TableRow()
        'tRow.Font.Bold = True
        'row=9 cell=1
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "待測板上下預留空間<br>電路板上下方零件高度"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=9 cell 2,3
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Left
        Labelx = New Label
        Labelx.ID = "label_9_2_1"
        Labelx.Text = "設備規格標準:上"
        tCell.Controls.Add(Labelx)
        tTxt = New TextBox()
        tTxt.ID = "txt_topspace"
        tTxt.Font.Name = "標楷體"
        tTxt.Font.Size = 14
        tTxt.Width = 50
        tTxt.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_9_2_2"
        Labelx.Text = "mm/下"
        tCell.Controls.Add(Labelx)
        tTxt = New TextBox()
        tTxt.ID = "txt_botspace"
        tTxt.Font.Name = "標楷體"
        tTxt.Font.Size = 14
        tTxt.Width = 50
        tTxt.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_9_2_3"
        Labelx.Text = "mm"
        tCell.Controls.Add(Labelx)
        tRow.Controls.Add(tCell)
        'row=9 cell=4
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "上下鏡組需求"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=9 cell=5,6
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Left
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_tblens"
        rRBL.Items.Add("上鏡組")
        rRBL.Items.Add("下鏡組")
        rRBL.Items.Add("上下鏡組")
        rRBL.RepeatDirection = RepeatDirection.Vertical
        rRBL.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(rRBL)
        cChk = New CheckBox
        cChk.ID = "chk_sidecamera"
        cChk.Text = "側面相機"
        tCell.Controls.Add(cChk)
        tRow.Controls.Add(tCell)

        ContentT.Rows.Add(tRow)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        tRow = New TableRow()
        'tRow.Font.Bold = True
        'row=10 cell=1
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "相機規格"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=10 cell=2,3
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Left
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_camerapixel"
        rRBL.Items.Add("6.5M")
        rRBL.Items.Add("12M")
        rRBL.Items.Add("21M")
        rRBL.Items.Add("25M")
        rRBL.Items.Add("37M")
        rRBL.Items.Add("其它")
        rRBL.RepeatDirection = RepeatDirection.Vertical
        rRBL.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(rRBL)
        tRow.Controls.Add(tCell)
        'row=10 cell=4
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "RGB光盤規格"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=10 cell=5,6
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Left
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_rgb"
        rRBL.Items.Add("自製(如V5/V7)")
        rRBL.Items.Add("外購(如OPT)")
        rRBL.RepeatDirection = RepeatDirection.Vertical
        rRBL.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(rRBL)
        tRow.Controls.Add(tCell)
        ContentT.Rows.Add(tRow)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        tRow = New TableRow()
        'tRow.Font.Bold = True
        'row=11 cell=1
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "系統解析度"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=11 cell=2-5
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 4
        tCell.HorizontalAlign = HorizontalAlign.Left
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_resolution"
        rRBL.Items.Add("20um")
        rRBL.Items.Add("15um")
        rRBL.Items.Add("12um")
        rRBL.Items.Add("10um")
        rRBL.Items.Add("8um")
        rRBL.Items.Add("7um")
        rRBL.Items.Add("6um")
        rRBL.Items.Add("5.5um")
        rRBL.Items.Add("5um")
        rRBL.Items.Add("3um")
        rRBL.Items.Add("2.5um")
        rRBL.Items.Add("其它")
        rRBL.RepeatDirection = RepeatDirection.Vertical
        rRBL.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(rRBL)
        tRow.Controls.Add(tCell)
        'row=11 cell=6
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        tTxt = New TextBox()
        tTxt.ID = "txt_otherresolution"
        tTxt.Font.Name = "標楷體"
        tTxt.Font.Size = 14
        tTxt.Width = 50
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_11_2"
        Labelx.Text = "&nbspum"
        tCell.Controls.Add(Labelx)
        tRow.Controls.Add(tCell)
        ContentT.Rows.Add(tRow)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        tRow = New TableRow()
        'tRow.Font.Bold = True
        'row=12 cell=1
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "光源控制器"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=12 cell=2,3
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Left
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_rbgcontrol"
        rRBL.Items.Add("自製(如V5/V7)")
        rRBL.Items.Add("外購(如OPT)")
        rRBL.RepeatDirection = RepeatDirection.Vertical
        rRBL.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(rRBL)
        tRow.Controls.Add(tCell)
        'row=12 cell=4
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "同軸光"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=12 cell=5
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Left
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_coaxialinstall"
        rRBL.Items.Add("加裝")
        rRBL.Items.Add("不加裝")
        rRBL.RepeatDirection = RepeatDirection.Horizontal
        rRBL.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(rRBL)
        tRow.Controls.Add(tCell)
        'row=12 cell=6
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Left
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_coaxialcolor"
        rRBL.Items.Add("紅")
        rRBL.Items.Add("白")
        rRBL.Items.Add("紅白")
        rRBL.RepeatDirection = RepeatDirection.Horizontal
        tCell.Controls.Add(rRBL)
        tRow.Controls.Add(tCell)
        ContentT.Rows.Add(tRow)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        tRow = New TableRow()
        'tRow.Font.Bold = True
        'row=13 cell=1
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "軌道皮帶型式與抗靜電需求"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=13 cell=2,3
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Left
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_belttype"
        rRBL.Items.Add("平面皮帶(產品重量 < 5 KG)")
        rRBL.Items.Add("時規皮帶(產品重量 > 5 KG)")
        rRBL.RepeatDirection = RepeatDirection.Horizontal
        rRBL.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(rRBL)
        cChk = New CheckBox
        cChk.ID = "chk_flux"
        cChk.Text = "考慮助焊劑沾黏問題"
        tCell.Controls.Add(cChk)
        cChk = New CheckBox
        cChk.ID = "chk_anti"
        cChk.Text = "具備抗靜電"
        tCell.Controls.Add(cChk)
        tRow.Controls.Add(tCell)
        'row=13 cell=4
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "軌道皮帶露出寬度(與電路板或載具的接處寬度"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=13 cell=5,6
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.Text = "設備規格為3.5mm以上"
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        ContentT.Rows.Add(tRow)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'row=14 cell=1-6
        tRow = New TableRow()
        'tRow.Font.Bold = True

        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.ColumnSpan = 6
        tCell.HorizontalAlign = HorizontalAlign.Left
        tCell.Text = "備註:"
        tCell.Font.Bold = True
        tRow.Controls.Add(tCell)
        ContentT.Rows.Add(tRow)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        tRow = New TableRow()
        'tRow.Font.Bold = True

        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Left
        tCell.ColumnSpan = 6
        tTxt = New TextBox()
        tTxt.ID = "txt_memo"
        tTxt.Width = 1250
        tTxt.Font.Name = "標楷體"
        tTxt.Font.Size = 14
        tTxt.TextMode = TextBoxMode.MultiLine
        tTxt.Rows = 6
        tTxt.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(tTxt)
        tRow.Controls.Add(tCell)
        ContentT.Rows.Add(tRow)
        'ShowXMSCT()
    End Sub
    Sub InitTableToSfid49_50_51_100()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim tablerow As Integer
        Dim Hyper As HyperLink
        Dim ChkMaterialDel As CheckBox
        Dim mcount As Integer
        Dim BColor As Drawing.Color
        Dim tImage As Image
        Dim colcountadjust As Integer
        Dim Labelx As Label
        colcountadjust = 0
        If (docstatus <> "B" And docstatus <> "E" And docstatus <> "D" And docstatus <> "R" And docstatus <> "A") Then '因送審後 , 有 2個欄位會刪除
            colcountadjust = 2
        End If
        BColor = System.Drawing.Color.LightBlue
        ContentT.Font.Name = "標楷體"
        ContentT.Font.Size = 12
        SqlCmd = "Select count(*) FROM [dbo].[@XSMLS] T0 WHERE head=0 and T0.[docentry] =" & docnum
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        drL.Read()
        mcount = drL(0)
        If (sfid <> 100) Then
            If (drL(0) <= 5) Then
                tablerow = 6
            Else
                tablerow = drL(0) + 1
            End If
        Else
            If (drL(0) < 2) Then
                tablerow = 2
            Else
                tablerow = drL(0) + 1
            End If
        End If
        drL.Close()
        connL.Close()
        'logo
        tRow = New TableRow()
        For j = 1 To (16 - colcountadjust) '10
            tCell = New TableCell
            tCell.BorderWidth = 0
            tCell.Width = 150
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Controls.Add(tCell)
        Next
        FormLogoTitleT.Rows.Add(tRow)
        FormLogoTitleT.Font.Name = "標楷體"
        FormLogoTitleT.Font.Size = 12
        tRow = New TableRow()
        tRow.Font.Bold = True
        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Left
        tCell.ColumnSpan = 3
        tImage = New Image
        tImage.ID = "image_logo"
        tImage.ImageUrl = "~/image/jetlog80%.jpg"
        tCell.Controls.Add(tImage)
        tRow.Controls.Add(tCell)
        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.Font.Size = 24
        tCell.ColumnSpan = 10 - colcountadjust
        tCell.HorizontalAlign = HorizontalAlign.Center
        If (sfid = 49) Then
            tCell.Text = "捷智科技 生產料件報廢單"
        ElseIf (sfid = 50) Then
            tCell.Text = "捷智科技 備品需求聯絡單"
        ElseIf (sfid = 51) Then
            tCell.Text = "捷智科技 料件入出庫單"
        ElseIf (sfid = 100) Then
            tCell.Text = "捷智科技 已簽核單據之補充單"
        End If
        tRow.Controls.Add(tCell)
        tCell = New TableCell
        tCell.Font.Size = 12
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Right
        tCell.VerticalAlign = VerticalAlign.Bottom
        tCell.ColumnSpan = 3
        If (docnum <> 0) Then
            If (docstatus = "E" Or docstatus = "D") Then
                SqlCmd = "Select convert(varchar(12), docdate, 111) from [dbo].[@XASCH] where docnum=" & docnum
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    tCell.Text = "單號:" & docnum & "<br>建單日期:" & drL(0)
                End If
                drL.Close()
                connL.Close()
            Else
                SqlCmd = "Select convert(varchar(12), signdate, 111) from [dbo].[@XSPWT] where docentry=" & docnum & " and seq=1"
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    tCell.Text = "單號:" & docnum & "<br>送審日期:" & drL(0)
                End If
                drL.Close()
                connL.Close()
            End If
        End If
        tRow.Controls.Add(tCell)
        FormLogoTitleT.Rows.Add(tRow)

        'tRow = New TableRow()
        'For j = 1 To 10
        '    tCell = New TableCell
        '    tCell.BorderWidth = 0
        '    tCell.Width = 200
        '    tCell.HorizontalAlign = HorizontalAlign.Center
        '    tRow.Controls.Add(tCell)
        'Next
        'ContentT.Rows.Add(tRow)
        '事由說明表頭列
        tRow = New TableRow()
        tRow.BackColor = Drawing.Color.LightBlue
        tRow.Font.Bold = True
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 16 - colcountadjust
        tCell.HorizontalAlign = HorizontalAlign.Center
        If (sfid <> 100) Then
            tCell.Text = "事由說明"
        Else
            tCell.Text = "補充說明"
        End If
        tRow.Controls.Add(tCell)
        tRow.Font.Size = 16
        ContentT.Rows.Add(tRow)
        '事由輸入列
        tRow = New TableRow()
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.Wrap = False
        tCell.HorizontalAlign = HorizontalAlign.Left
        tCell.ColumnSpan = 16 - colcountadjust
        TxtReason = New TextBox
        TxtReason.ID = "txt_reason"
        TxtReason.TextMode = TextBoxMode.MultiLine
        TxtReason.Rows = 7
        TxtReason.Width = 1000
        'AddHandler TxtReason.TextChanged, AddressOf TxtReason_TextChanged
        'TxtReason.AutoPostBack = True
        'TxtReason.BackColor = Drawing.Color.Cornsilk
        tCell.Controls.Add(TxtReason)
        tRow.Cells.Add(tCell)
        ContentT.Rows.Add(tRow)
        '料件表頭列
        tRow = New TableRow()
        tRow.Font.Size = 16
        tRow.BackColor = Drawing.Color.LightBlue
        tRow.Font.Bold = True
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 16 - colcountadjust
        tCell.HorizontalAlign = HorizontalAlign.Center
        Labelx = New Label
        Labelx.ID = "label_title"
        If (sfid = 100) Then
            If (docstatus <> "B" And docstatus <> "E" And docstatus <> "D" And docstatus <> "R" And docstatus <> "A") Then
                Labelx.Text = "補充加簽之增減料件表列"
            Else
                Labelx.Text = "補充加簽之增減料件表列(請由上述橘色功能列加入)"
            End If
        Else
            If (docstatus <> "B" And docstatus <> "E" And docstatus <> "D" And docstatus <> "R" And docstatus <> "A") Then
                Labelx.Text = "增減料件表列"
            Else
                Labelx.Text = "增減料件表列(請由上述橘色功能列加入)"
            End If
        End If
        tCell.Controls.Add(Labelx)
        tRow.Controls.Add(tCell) 'hello
        ContentT.Rows.Add(tRow)

        tRow = New TableRow()
        tRow.BackColor = Drawing.Color.LightGreen
        tRow.Font.Bold = True
        For i = 0 To 15 - colcountadjust
            tCell = New TableCell
            tCell.BorderWidth = 1
            tCell.Width = 40
            tCell.HorizontalAlign = HorizontalAlign.Center
            If (i = 0) Then
                tCell.Text = "項次"
                tCell.Width = 40
            ElseIf (i = 1) Then
                tCell.Text = "料號"
                tCell.Width = 120
            ElseIf (i = 2) Then
                tCell.Text = "說明"
                tCell.Width = 300
            ElseIf (i = 3) Then
                tCell.Text = "需求"
                tCell.Width = 40
            ElseIf (i = 4) Then
                tCell.Text = "單價"
                tCell.Width = 40
            ElseIf (i = 5) Then
                tCell.Text = "總價"
                tCell.Width = 40
                '-------------------------------------
            ElseIf (i = 6) Then
                tCell.Text = "本庫"
                tCell.Width = 40
            ElseIf (i = 7) Then
                tCell.Text = "本需"
                tCell.Width = 40
            ElseIf (i = 8) Then
                tCell.Text = "本供"
                tCell.Width = 40
            ElseIf (i = 9) Then
                tCell.Text = "它庫"
                tCell.Width = 40
            ElseIf (i = 10) Then
                tCell.Text = "它需"
                tCell.Width = 40
            ElseIf (i = 11) Then
                tCell.Text = "它供"
                tCell.Width = 40
            ElseIf (i = 12) Then '6
                If (sfid = 51 Or sfid = 50 Or sfid = 100) Then
                    tCell.Text = "處置"
                ElseIf (sfid = 49) Then
                    tCell.Text = "報廢原因"
                End If
                tCell.Width = 160
            ElseIf (i = 13) Then
                tCell.Text = "備註"
                tCell.Width = 250
            ElseIf (i = 14) Then
                tCell.Text = "動作"
                tCell.Width = 50
                'If (docstatus <> "B" And docstatus <> "E" And docstatus <> "D" And docstatus <> "R" And docstatus <> "A") Then
                'tCell.Visible = False
                'End If
            ElseIf (i = 15) Then
                tCell.Text = "刪除"
                tCell.Width = 60
                'If (docstatus <> "B" And docstatus <> "E" And docstatus <> "D" And docstatus <> "R" And docstatus <> "A") Then
                'tCell.Visible = False
                'End If
            End If
            tRow.Controls.Add(tCell)
        Next
        ContentT.Rows.Add(tRow)
        For i = 4 To tablerow + 3 '原 i=1 to tablerow , 但此之前有幾列row , 故需加上 , 以便用i命名之id 能與showmaterial 一致
            tRow = New TableRow()
            tRow.BorderWidth = 1
            For j = 0 To 15 - colcountadjust
                tCell = New TableCell
                tCell.BorderWidth = 1
                tCell.Height = 20
                If (j = 0 Or j = 3 Or j = 4 Or j = 5 Or j = 14 Or j = 15 Or j = 6 Or j = 7 Or j = 8 Or j = 9 Or j = 10 Or j = 11) Then
                    tCell.HorizontalAlign = HorizontalAlign.Center
                End If
                If (i <= mcount + 3) Then '加3同上說明
                    If (j = 14) Then
                        Hyper = New HyperLink()
                        Hyper.ID = "hypermodify_" & i & "_8"
                        Hyper.Text = "修改"
                        'If (docstatus <> "B" And docstatus <> "E" And docstatus <> "D" And docstatus <> "R" And docstatus <> "A") Then
                        'tCell.Visible = False
                        'End If
                        tCell.Controls.Add(Hyper)
                    End If
                    If (j = 15) Then
                        ChkMaterialDel = New CheckBox
                        ChkMaterialDel.ID = "chkdel_" & i & "_9"
                        ChkMaterialDel.Text = "刪除"
                        ChkMaterialDel.AutoPostBack = True
                        'If (docstatus <> "B" And docstatus <> "E" And docstatus <> "D" And docstatus <> "R" And docstatus <> "A") Then
                        'tCell.Visible = False
                        'End If
                        AddHandler ChkMaterialDel.CheckedChanged, AddressOf ChkMaterialDel_CheckedChanged
                        tCell.Controls.Add(ChkMaterialDel)
                    End If
                Else
                    If (j = 14 Or j = 15) Then
                        'If (docstatus <> "B" And docstatus <> "E" And docstatus <> "D" And docstatus <> "R" And docstatus <> "A") Then
                        'tCell.Visible = False
                        'End If
                    End If
                End If
                tRow.Controls.Add(tCell)
            Next
            ContentT.Rows.Add(tRow)
        Next
    End Sub
    Sub InitTableToSfid23_24()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim tablerow As Integer
        Dim Hyper As HyperLink
        Dim ChkMaterialDel As CheckBox
        Dim mcount As Integer
        Dim BColor As Drawing.Color
        Dim tImage As Image
        Dim colcountadjust As Integer
        colcountadjust = 0
        If (docstatus <> "B" And docstatus <> "E" And docstatus <> "D" And docstatus <> "R" And docstatus <> "A") Then '因送審後 , 有 2個欄位會刪除
            colcountadjust = 2
        End If
        BColor = System.Drawing.Color.LightBlue
        ContentT.Font.Name = "標楷體"
        ContentT.Font.Size = 12
        SqlCmd = "Select count(*) FROM [dbo].[@XSMLS] T0 WHERE head=0 and T0.[docentry] =" & docnum
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        drL.Read()
        mcount = drL(0)
        If (drL(0) <= 5) Then
            tablerow = 6
        Else
            tablerow = drL(0) + 1
        End If
        drL.Close()
        connL.Close()
        'logo
        tRow = New TableRow()
        For j = 1 To (9 - colcountadjust) '10
            tCell = New TableCell
            tCell.BorderWidth = 0
            tCell.Width = 300
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Controls.Add(tCell)
        Next
        FormLogoTitleT.Rows.Add(tRow)
        FormLogoTitleT.Font.Name = "標楷體"
        FormLogoTitleT.Font.Size = 12
        tRow = New TableRow()
        tRow.Font.Bold = True
        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Left
        'tCell.ColumnSpan = 3
        tImage = New Image
        tImage.ID = "image_logo"
        tImage.ImageUrl = "~/image/jetlog80%.jpg"
        tCell.Controls.Add(tImage)
        tRow.Controls.Add(tCell)
        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.Font.Size = 24
        tCell.ColumnSpan = 7 - colcountadjust
        tCell.HorizontalAlign = HorizontalAlign.Center
        If (sfid = 23) Then
            tCell.Text = "捷智科技 離倉料件管制單"
        ElseIf (sfid = 24) Then
            tCell.Text = "捷智科技 借入料件管制單"
        End If
        tRow.Controls.Add(tCell)
        tCell = New TableCell
        tCell.Font.Size = 12
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Right
        tCell.VerticalAlign = VerticalAlign.Bottom
        'tCell.ColumnSpan = 3
        If (docnum <> 0) Then
            If (docstatus = "E" Or docstatus = "D") Then
                SqlCmd = "Select convert(varchar(12), docdate, 111) from [dbo].[@XASCH] where docnum=" & docnum
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    tCell.Text = "單號:" & docnum & "<br>建單日期:" & drL(0)
                End If
                drL.Close()
                connL.Close()
            Else
                SqlCmd = "Select convert(varchar(12), signdate, 111) from [dbo].[@XSPWT] where docentry=" & docnum & " and seq=1"
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    tCell.Text = "單號:" & docnum & "<br>送審日期:" & drL(0)
                End If
                drL.Close()
                connL.Close()
            End If
        End If
        tRow.Controls.Add(tCell)
        FormLogoTitleT.Rows.Add(tRow)

        '事由說明表頭列
        tRow = New TableRow()
        tRow.BackColor = Drawing.Color.LightBlue
        tRow.Font.Bold = True
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 9 - colcountadjust
        tCell.HorizontalAlign = HorizontalAlign.Center
        If (sfid = 23) Then
            tCell.Text = "離倉料件說明"
        ElseIf (sfid = 24) Then
            tCell.Text = "借入料件說明"
        End If
        tRow.Controls.Add(tCell)
        tRow.Font.Size = 16
        ContentT.Rows.Add(tRow)
        '事由輸入列
        tRow = New TableRow()
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.Wrap = False
        tCell.HorizontalAlign = HorizontalAlign.Left
        tCell.ColumnSpan = 9 - colcountadjust
        TxtReason = New TextBox
        TxtReason.ID = "txt_reason"
        TxtReason.TextMode = TextBoxMode.MultiLine
        TxtReason.Rows = 7
        TxtReason.Width = 1000
        'AddHandler TxtReason.TextChanged, AddressOf TxtReason_TextChanged
        'TxtReason.AutoPostBack = True
        'TxtReason.BackColor = Drawing.Color.Cornsilk
        tCell.Controls.Add(TxtReason)
        tRow.Cells.Add(tCell)
        ContentT.Rows.Add(tRow)
        '料件表頭列
        tRow = New TableRow()
        tRow.Font.Size = 16
        tRow.BackColor = Drawing.Color.LightBlue
        tRow.Font.Bold = True
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 9 - colcountadjust
        tCell.HorizontalAlign = HorizontalAlign.Center
        If (sfid = 23) Then
            If (docstatus <> "B" And docstatus <> "E" And docstatus <> "D" And docstatus <> "R" And docstatus <> "A") Then
                tCell.Text = "離倉料件表列"
            Else
                tCell.Text = "離倉料件表列(請由上述橘色功能列加入)"
            End If
        ElseIf (sfid = 24) Then
            If (docstatus <> "B" And docstatus <> "E" And docstatus <> "D" And docstatus <> "R" And docstatus <> "A") Then
                tCell.Text = "借入料件表列"
            Else
                tCell.Text = "借入料件表列(請由上述橘色功能列加入)"
            End If
        End If
        tRow.Controls.Add(tCell)
        ContentT.Rows.Add(tRow)

        tRow = New TableRow()
        tRow.BackColor = Drawing.Color.LightGreen
        tRow.Font.Bold = True
        For i = 0 To 8 - colcountadjust
            tCell = New TableCell
            tCell.BorderWidth = 1
            tCell.Width = 40
            tCell.HorizontalAlign = HorizontalAlign.Center
            If (i = 0) Then
                tCell.Text = "項次"
                tCell.Width = 40
            ElseIf (i = 1) Then
                tCell.Text = "料號"
                tCell.Width = 200 '120
            ElseIf (i = 2) Then
                tCell.Text = "說明"
                tCell.Width = 500 '300
            ElseIf (i = 3) Then
                If (sfid = 23) Then
                    tCell.Text = "離倉數量"
                ElseIf (sfid = 24) Then
                    tCell.Text = "借入數量"
                End If
                tCell.Width = 40
            ElseIf (i = 4) Then
                tCell.Text = "已還數量"
                tCell.Width = 40
            ElseIf (i = 5) Then '6
                If (sfid = 23) Then
                    tCell.Text = "離倉原因"
                ElseIf (sfid = 24) Then
                    tCell.Text = "借入原因"
                End If
                tCell.Width = 200 '160
            ElseIf (i = 6) Then
                tCell.Text = "備註"
                tCell.Width = 300 '250
            ElseIf (i = 7) Then
                tCell.Text = "動作"
                tCell.Width = 50
                'If (docstatus <> "B" And docstatus <> "E" And docstatus <> "D" And docstatus <> "R" And docstatus <> "A") Then
                'tCell.Visible = False
                'End If
            ElseIf (i = 8) Then
                tCell.Text = "刪除"
                tCell.Width = 60
                'If (docstatus <> "B" And docstatus <> "E" And docstatus <> "D" And docstatus <> "R" And docstatus <> "A") Then
                'tCell.Visible = False
                'End If
            End If
            tRow.Controls.Add(tCell)
        Next
        ContentT.Rows.Add(tRow)
        For i = 4 To tablerow + 3 '原 i=1 to tablerow , 但此之前有幾列row , 故需加上 , 以便用i命名之id 能與showmaterial 一致
            tRow = New TableRow()
            tRow.BorderWidth = 1
            For j = 0 To 8 - colcountadjust
                tCell = New TableCell
                tCell.BorderWidth = 1
                tCell.Height = 20
                If (j = 0 Or j = 3 Or j = 4 Or j = 6 Or j = 7 Or j = 8) Then
                    tCell.HorizontalAlign = HorizontalAlign.Center
                End If
                If (i <= mcount + 3) Then '加3同上說明
                    If (j = 7) Then
                        Hyper = New HyperLink()
                        Hyper.ID = "hypermodify_" & i & "_8"
                        Hyper.Text = "修改"
                        'If (docstatus <> "B" And docstatus <> "E" And docstatus <> "D" And docstatus <> "R" And docstatus <> "A") Then
                        'tCell.Visible = False
                        'End If
                        tCell.Controls.Add(Hyper)
                    End If
                    If (j = 8) Then
                        ChkMaterialDel = New CheckBox
                        ChkMaterialDel.ID = "chkdel_" & i & "_9"
                        ChkMaterialDel.Text = "刪除"
                        ChkMaterialDel.AutoPostBack = True
                        'If (docstatus <> "B" And docstatus <> "E" And docstatus <> "D" And docstatus <> "R" And docstatus <> "A") Then
                        'tCell.Visible = False
                        'End If
                        AddHandler ChkMaterialDel.CheckedChanged, AddressOf ChkMaterialDel_CheckedChanged
                        tCell.Controls.Add(ChkMaterialDel)
                    End If
                End If
                tRow.Controls.Add(tCell)
            Next
            ContentT.Rows.Add(tRow)
        Next
    End Sub
    Sub InitTableToSfid101()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim tablerow As Integer
        'Dim Hyper As HyperLink
        'Dim ChkMaterialDel As CheckBox
        Dim mcount, i As Integer
        Dim BColor As Drawing.Color
        Dim tImage As Image
        Dim colcountadjust As Integer
        'Dim dDDL As DropDownList
        Dim tTxt As TextBox
        Dim mainsfid As Integer
        mainsfid = 0
        colcountadjust = 0
        'If (docstatus <> "B" And docstatus <> "E" And docstatus <> "D" And docstatus <> "R" And docstatus <> "A") Then '因送審後 , 有 2個欄位會刪除
        '    colcountadjust = 2
        'End If
        BColor = System.Drawing.Color.LightBlue
        ContentT.Font.Name = "標楷體"
        ContentT.Font.Size = 12
        If (TxtAttaDoc.Text <> "NA" And TxtAttaDoc.Text <> "") Then
            SqlCmd = "Select count(*) FROM [dbo].[@XSMLS] T0 WHERE head=0 and T0.[docentry] =" & CLng(TxtAttaDoc.Text)
            drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
            drL.Read()
            mcount = drL(0)
            If (drL(0) <= 5) Then
                tablerow = 6
            Else
                tablerow = drL(0) + 1
            End If
            drL.Close()
            connL.Close()
        Else
            tablerow = 6
        End If
        'logo
        tRow = New TableRow()
        For j = 1 To (8 - colcountadjust) '10
            tCell = New TableCell
            tCell.BorderWidth = 0
            tCell.Width = 300
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Controls.Add(tCell)
        Next
        FormLogoTitleT.Rows.Add(tRow)
        FormLogoTitleT.Font.Name = "標楷體"
        FormLogoTitleT.Font.Size = 12
        tRow = New TableRow()
        tRow.Font.Bold = True
        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Left
        'tCell.ColumnSpan = 3
        tImage = New Image
        tImage.ID = "image_logo"
        tImage.ImageUrl = "~/image/jetlog80%.jpg"
        tCell.Controls.Add(tImage)
        tRow.Controls.Add(tCell)
        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.Font.Size = 24
        tCell.ColumnSpan = 6 - colcountadjust
        tCell.HorizontalAlign = HorizontalAlign.Center
        If (TxtAttaDoc.Text <> "NA" And TxtAttaDoc.Text <> "") Then
            SqlCmd = "Select sfid from [dbo].[@XASCH] where docnum=" & CLng(TxtAttaDoc.Text)
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                dr.Read()
                mainsfid = dr(0)
            End If
            dr.Close()
            connsap.Close()
            If (mainsfid = 23) Then
                tCell.Text = "捷智科技 料件返還單(對應離倉單號:" & TxtAttaDoc.Text & ")"
            Else
                tCell.Text = "捷智科技 料件返還單(對應借入單號:" & TxtAttaDoc.Text & ")"
            End If
        Else
            tCell.Text = "捷智科技 料件返還單"
        End If
        tRow.Controls.Add(tCell)
        tCell = New TableCell
        tCell.Font.Size = 12
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Right
        tCell.VerticalAlign = VerticalAlign.Bottom
        'tCell.ColumnSpan = 3
        If (docnum <> 0) Then
            If (docstatus = "E" Or docstatus = "D") Then
                SqlCmd = "Select convert(varchar(12), docdate, 111) from [dbo].[@XASCH] where docnum=" & docnum
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    tCell.Text = "單號:" & docnum & "<br>建單日期:" & drL(0)
                End If
                drL.Close()
                connL.Close()
            Else
                SqlCmd = "Select convert(varchar(12), signdate, 111) from [dbo].[@XSPWT] where docentry=" & docnum & " and seq=1"
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    tCell.Text = "單號:" & docnum & "<br>送審日期:" & drL(0)
                End If
                drL.Close()
                connL.Close()
            End If
        End If
        tRow.Controls.Add(tCell)
        FormLogoTitleT.Rows.Add(tRow)

        ''事由說明表頭列
        'tRow = New TableRow()
        'tRow.BackColor = Drawing.Color.LightBlue
        'tRow.Font.Bold = True
        'tCell = New TableCell
        'tCell.BorderWidth = 1
        'tCell.ColumnSpan = 16 - colcountadjust
        'tCell.HorizontalAlign = HorizontalAlign.Center
        'If (sfid = 23) Then
        '    tCell.Text = "料件還回說明"
        'End If
        'tRow.Controls.Add(tCell)
        'tRow.Font.Size = 16
        'ContentT.Rows.Add(tRow)
        ''事由輸入列
        'tRow = New TableRow()
        'tCell = New TableCell()
        'tCell.BorderWidth = 1
        'tCell.Wrap = False
        'tCell.HorizontalAlign = HorizontalAlign.Left
        'tCell.ColumnSpan = 16 - colcountadjust
        'TxtReason = New TextBox
        'TxtReason.ID = "txt_reason"
        'TxtReason.TextMode = TextBoxMode.MultiLine
        'TxtReason.Rows = 7
        'TxtReason.Width = 1000
        ''AddHandler TxtReason.TextChanged, AddressOf TxtReason_TextChanged
        ''TxtReason.AutoPostBack = True
        ''TxtReason.BackColor = Drawing.Color.Cornsilk
        'tCell.Controls.Add(TxtReason)
        'tRow.Cells.Add(tCell)
        'ContentT.Rows.Add(tRow)
        '料件表頭列
        tRow = New TableRow()
        tRow.Font.Size = 16
        tRow.BackColor = Drawing.Color.LightBlue
        tRow.Font.Bold = True
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 8 - colcountadjust
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "料件返還表列"
        tRow.Controls.Add(tCell)
        ContentT.Rows.Add(tRow)

        tRow = New TableRow()
        tRow.BackColor = Drawing.Color.LightGreen
        tRow.Font.Bold = True
        For i = 0 To 7 - colcountadjust
            tCell = New TableCell
            tCell.BorderWidth = 1
            tCell.Width = 40
            tCell.HorizontalAlign = HorizontalAlign.Center
            If (i = 0) Then
                tCell.Text = "項次"
                tCell.Width = 40
            ElseIf (i = 1) Then
                tCell.Text = "料號"
                tCell.Width = 200 '120
            ElseIf (i = 2) Then
                tCell.Text = "說明"
                tCell.Width = 500 '300
            ElseIf (i = 3) Then
                If (mainsfid = 23) Then
                    tCell.Text = "離倉數量"
                ElseIf (mainsfid = 24) Then
                    tCell.Text = "借入數量"
                Else
                    tCell.Text = "數量"
                End If
                tCell.Width = 40
            ElseIf (i = 4) Then
                tCell.Text = "已還數量"
                tCell.Width = 40
            ElseIf (i = 5) Then
                tCell.Text = "此次返還"
                tCell.Width = 40
            ElseIf (i = 6) Then
                If (mainsfid = 23) Then
                    tCell.Text = "當初離倉原因"
                ElseIf (mainsfid = 24) Then
                    tCell.Text = "當初借入原因"
                Else
                    tCell.Text = "當初原因"
                End If
                tCell.Width = 40
            ElseIf (i = 7) Then
                tCell.Text = "備註"
                tCell.Width = 300 '250
            End If
            tRow.Controls.Add(tCell)
        Next
        ContentT.Rows.Add(tRow)
        For i = 2 To tablerow + 1 '原 i=1 to tablerow , 但此之前有幾列row , 故需加上 , 以便用i命名之id 能與showmaterial 一致
            tRow = New TableRow()
            tRow.BorderWidth = 1
            For j = 0 To 7 - colcountadjust
                tCell = New TableCell
                tCell.BorderWidth = 1
                tCell.Height = 20
                If (j = 5) Then
                    tCell.Width = 40
                ElseIf (j = 7) Then
                    tCell.Width = 200
                End If
                If (j = 0 Or j = 3 Or j = 4 Or j = 5 Or j = 6) Then
                    tCell.HorizontalAlign = HorizontalAlign.Center
                End If
                tRow.Controls.Add(tCell)
            Next
            ContentT.Rows.Add(tRow)
        Next
        If (TxtAttaDoc.Text <> "NA" And TxtAttaDoc.Text <> "") Then
            i = 2 '料件在Table 之起始列
            SqlCmd = "Select T0.itemcode,T0.itemname,T0.quantity,T0.rtnqty,T0.method,T0.comment,T0.num,T0.nowrtnqty " &
                "FROM [dbo].[@XSMLS] T0 " &
                "WHERE T0.head=0 And T0.[docentry] =" & CLng(TxtAttaDoc.Text) & " ORDER BY T0.num"
            drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
            If (drL.HasRows) Then
                Do While (drL.Read())
                    ContentT.Rows(i).Cells(0).Text = i - 1 '項次
                    ContentT.Rows(i).Cells(1).Text = drL(0) '料號
                    ContentT.Rows(i).Cells(2).Text = drL(1) '說明
                    ContentT.Rows(i).Cells(3).Text = drL(2) '離倉數量
                    ContentT.Rows(i).Cells(4).Text = drL(3) '已還數量
                    If (docstatus = "E" Or docstatus = "D" Or docstatus = "R" Or docstatus = "B") Then
                        tTxt = New TextBox
                        tTxt.ID = "txtreturn_" & i & "_" & drL(6)
                        tTxt.Width = 40
                        'tTxt.Enabled = False
                        tTxt.Text = drL(7)
                        tTxt.AutoPostBack = True
                        AddHandler tTxt.TextChanged, AddressOf tTxtreturn_TextChanged
                        ContentT.Rows(i).Cells(5).Controls.Add(tTxt)
                    Else
                        ContentT.Rows(i).Cells(5).Text = drL(7)
                    End If

                    ContentT.Rows(i).Cells(6).Text = drL(4) '當初離倉原因
                    If (docstatus = "E" Or docstatus = "D" Or docstatus = "R" Or docstatus = "B") Then
                        tTxt = New TextBox
                        tTxt.ID = "txtnote_" & i & "_" & drL(6)
                        tTxt.Width = 200
                        'tTxt.Enabled = False
                        tTxt.Text = drL(5)
                        tTxt.AutoPostBack = True
                        AddHandler tTxt.TextChanged, AddressOf tTxtNote_TextChanged
                        ContentT.Rows(i).Cells(7).Controls.Add(tTxt)
                    Else
                        ContentT.Rows(i).Cells(7).Text = drL(5) '備註
                    End If
                    i = i + 1 'qqqqq
                Loop
            End If
            drL.Close()
            connL.Close()
        End If
    End Sub
    Sub InitTableToSfid1()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim BColor As Drawing.Color
        Dim tImage As Image
        tRow = New TableRow()
        For j = 1 To 10
            tCell = New TableCell
            tCell.BorderWidth = 0
            tCell.Width = 200
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Controls.Add(tCell)
        Next
        FormLogoTitleT.Rows.Add(tRow) 'row=0
        BColor = System.Drawing.Color.LightBlue
        FormLogoTitleT.Font.Name = "標楷體"
        FormLogoTitleT.Font.Size = 12
        tRow = New TableRow()
        tRow.Font.Bold = True
        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Left
        tCell.ColumnSpan = 1
        tImage = New Image
        tImage.ID = "image_logo"
        tImage.ImageUrl = "~/image/jetlog80%.jpg"
        tCell.Controls.Add(tImage)
        tRow.Controls.Add(tCell)
        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.Font.Size = 24
        tCell.ColumnSpan = 8
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "捷智科技 內部聯絡單"
        tRow.Controls.Add(tCell)
        tCell = New TableCell
        tCell.Font.Size = 12
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Right
        tCell.VerticalAlign = VerticalAlign.Bottom
        tCell.ColumnSpan = 1
        If (docnum <> 0) Then
            If (docstatus = "E" Or docstatus = "D") Then
                SqlCmd = "Select convert(varchar(12), docdate, 111) from [dbo].[@XASCH] where docnum=" & docnum
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    tCell.Text = "單號:" & docnum & "<br>建單日期:" & drL(0)
                End If
                drL.Close()
                connL.Close()
            Else
                SqlCmd = "Select convert(varchar(12), signdate, 111) from [dbo].[@XSPWT] where docentry=" & docnum & " and seq=1"
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    tCell.Text = "單號:" & docnum & "<br>送審日期:" & drL(0)
                End If
                drL.Close()
                connL.Close()
            End If
        End If
        tRow.Controls.Add(tCell)
        FormLogoTitleT.Rows.Add(tRow) 'row=1

        tRow = New TableRow()
        For j = 1 To 10
            tCell = New TableCell
            tCell.BorderWidth = 0
            tCell.Width = 200
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Controls.Add(tCell)
        Next
        ContentT.Rows.Add(tRow) 'row=0
        BColor = System.Drawing.Color.LightBlue
        ContentT.Font.Name = "標楷體"
        ContentT.Font.Size = 14
        '聯絡表頭 
        tRow = New TableRow()
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "送審日期"
        tCell.BackColor = Drawing.Color.LightBlue
        tRow.Controls.Add(tCell)

        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.Wrap = False
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.ColumnSpan = 3
        tRow.Cells.Add(tCell)

        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "送審人"
        tCell.BackColor = Drawing.Color.LightBlue
        tRow.Controls.Add(tCell)

        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.Wrap = False
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.ColumnSpan = 3
        tRow.Cells.Add(tCell)
        ContentT.Rows.Add(tRow) 'row=1
        '第二列
        tRow = New TableRow()
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "聯絡部門"
        tCell.BackColor = Drawing.Color.LightBlue
        tRow.Controls.Add(tCell)

        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.Wrap = False
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.ColumnSpan = 3
        TxtDept = New TextBox
        TxtDept.ID = "txt_dept"
        TxtDept.Font.Size = 14
        TxtDept.Width = 400
        tCell.Controls.Add(TxtDept)
        tRow.Cells.Add(tCell)

        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "聯絡人員"
        tCell.BackColor = Drawing.Color.LightBlue
        tRow.Controls.Add(tCell)

        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.Wrap = False
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.ColumnSpan = 3
        TxtPerson = New TextBox
        TxtPerson.ID = "txt_person"
        TxtPerson.Width = 400
        TxtPerson.Font.Size = 14
        tCell.Controls.Add(TxtPerson)
        tRow.Cells.Add(tCell)
        ContentT.Rows.Add(tRow) 'row=2
        '主旨列
        tRow = New TableRow()
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "主旨"
        tCell.BackColor = Drawing.Color.LightBlue
        tRow.Controls.Add(tCell)

        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 8
        tCell.HorizontalAlign = HorizontalAlign.Left
        tRow.Controls.Add(tCell)
        ContentT.Rows.Add(tRow) 'row=3
        '事由說明表頭列
        tRow = New TableRow()
        tRow.BackColor = Drawing.Color.LightBlue
        tRow.Font.Bold = True
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 10
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "事由說明"
        tRow.Controls.Add(tCell)
        ContentT.Rows.Add(tRow) 'row=4
        '事由輸入列
        tRow = New TableRow()
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.Wrap = False
        tCell.HorizontalAlign = HorizontalAlign.Left
        tCell.ColumnSpan = 10
        TxtReason = New TextBox
        TxtReason.ID = "txt_reason"
        TxtReason.TextMode = TextBoxMode.MultiLine
        TxtReason.Rows = 7
        TxtReason.Font.Size = 14
        TxtReason.Width = 1200
        'TxtReason.BackColor = Drawing.Color.Cornsilk
        tCell.Controls.Add(TxtReason)
        tRow.Cells.Add(tCell)
        ContentT.Rows.Add(tRow) 'row=5

    End Sub
    Sub InitTableToSfid3_22()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim BColor, NeedInputColor, WhiteColor As Drawing.Color
        Dim tImage As Image
        Dim tChk As CheckBox
        NeedInputColor = Drawing.Color.AntiqueWhite
        WhiteColor = Drawing.Color.White
        tRow = New TableRow()
        For j = 1 To 16
            tCell = New TableCell
            tCell.BorderWidth = 0
            tCell.Width = 200
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Controls.Add(tCell)
        Next
        FormLogoTitleT.Rows.Add(tRow) 'row=0
        BColor = System.Drawing.Color.LightBlue
        FormLogoTitleT.Font.Name = "標楷體"
        FormLogoTitleT.Font.Size = 12
        tRow = New TableRow()
        tRow.Font.Bold = True
        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Left
        tCell.ColumnSpan = 1
        tImage = New Image
        tImage.ID = "image_logo"
        tImage.ImageUrl = "~/image/jetlog80%.jpg"
        tCell.Controls.Add(tImage)
        tRow.Controls.Add(tCell)
        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.Font.Size = 24
        tCell.ColumnSpan = 14
        tCell.HorizontalAlign = HorizontalAlign.Center
        If (sfid = 22) Then
            tCell.Text = "捷智科技 客戶機台問題反應單"
        ElseIf (sfid = 3) Then
            tCell.Text = "捷智科技 廠內機台問題反應單"
        End If
        tRow.Controls.Add(tCell)
        tCell = New TableCell
        tCell.Font.Size = 12
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Right
        tCell.VerticalAlign = VerticalAlign.Bottom
        tCell.ColumnSpan = 1
        If (docnum <> 0) Then
            If (docstatus = "E" Or docstatus = "D") Then
                SqlCmd = "Select convert(varchar(12), docdate, 111) from [dbo].[@XASCH] where docnum=" & docnum
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    tCell.Text = "單號:" & docnum & "<br>建單日期:" & drL(0)
                End If
                drL.Close()
                connL.Close()
            Else
                SqlCmd = "Select convert(varchar(12), signdate, 111) from [dbo].[@XSPWT] where docentry=" & docnum & " and seq=1"
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    tCell.Text = "單號:" & docnum & "<br>送審日期:" & drL(0)
                End If
                drL.Close()
                connL.Close()
            End If
        End If
        tRow.Controls.Add(tCell)
        FormLogoTitleT.Rows.Add(tRow) 'row=1

        tRow = New TableRow()
        For j = 1 To 16
            tCell = New TableCell
            tCell.BorderWidth = 0
            tCell.Width = 200
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Controls.Add(tCell)
        Next
        ContentT.Rows.Add(tRow) 'row=0
        BColor = System.Drawing.Color.LightBlue
        ContentT.Font.Name = "標楷體"
        ContentT.Font.Size = 14
        tRow = New TableRow()
        tRow.Font.Size = 16
        tRow.Controls.Add(CommUtil.CellSet("機台基本資訊", 1, 16, False, 0, 0, "center", BColor))
        ContentT.Rows.Add(tRow)  'row=1
        '聯絡表頭 
        tRow = New TableRow()
        'CellSet(Text As String, rowspan As Integer, colspan As Integer, FondBold As Boolean, txtid As String, width As Integer, height As Integer, align As String)
        tRow.Controls.Add(CommUtil.CellSet("回報日期", 1, 2, False, 0, 0, "center", BColor))
        tRow.Cells.Add(CommUtil.CellSetWithCalenderExtender(1, 2, "txt_reportdate", "ce_reportdate", NeedInputColor, 0))

        tRow.Controls.Add(CommUtil.CellSet("產品別", 1, 2, False, 0, 0, "center", BColor))
        'tRow.Controls.Add(CellSet("", 1, 2, False, 0, 0, "center", False))
        tRow.Controls.Add(CellSetWithExtender(1, 2, "lb_machinetype", "txt_machinetype", "dde_machinetype", NeedInputColor))

        tRow.Controls.Add(CommUtil.CellSet("客戶名稱", 1, 2, False, 0, 0, "center", BColor))
        tRow.Controls.Add(CellSetWithExtender(1, 2, "lb_cusname", "txt_cusname", "dde_cusname", NeedInputColor))

        If (sfid = 22) Then
            tRow.Controls.Add(CommUtil.CellSet("客戶廠區", 1, 2, False, 0, 0, "center", BColor))
            tRow.Controls.Add(CellSetWithExtender(1, 2, "lb_cusfactoryOrmo", "txt_cusfactoryOrmo", "dde_cusfactory", WhiteColor))
        ElseIf (sfid = 3) Then
            tRow.Controls.Add(CommUtil.CellSet("機台料號", 1, 2, False, 0, 0, "center", BColor))
            tRow.Cells.Add(CommUtil.CellSetWithTextBox(1, 2, "lb_cusfactoryOrmo", 0, 0, 0, NeedInputColor, "center"))
        End If
        ContentT.Rows.Add(tRow) 'row=2
        '第三列
        tRow = New TableRow()
        tRow.Controls.Add(CommUtil.CellSet("機台型號", 1, 2, False, 0, 0, "center", BColor))
        tRow.Controls.Add(CellSetWithExtender(1, 2, "lb_model", "txt_model", "dde_model", NeedInputColor))

        If (sfid = 22) Then
            tRow.Controls.Add(CommUtil.CellSet("機台序號", 1, 2, False, 0, 0, "center", BColor))
            tRow.Cells.Add(CommUtil.CellSetWithTextBox(1, 2, "txt_machineserialOrwo", 0, 0, 0, NeedInputColor, "center"))

            tRow.Controls.Add(CommUtil.CellSet("裝機日期", 1, 2, False, 0, 0, "center", BColor))
            tRow.Cells.Add(CommUtil.CellSetWithCalenderExtender(1, 2, "txt_installdateOrshipdate", "ce_installdate", WhiteColor, 0))

            'tCell = New TableCell
            tCell = CommUtil.CellSet("", 1, 4, False, 0, 0, "center", WhiteColor)
            tChk = New CheckBox
            tChk.ID = "chk_firstinstallOrnoassign"
            tChk.Text = "新安裝&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
            tCell.Controls.Add(tChk)
            tChk = New CheckBox
            tChk.ID = "chk_inwarranty"
            tChk.Text = "保固期內"
            tCell.Controls.Add(tChk)
            tRow.Controls.Add(tCell)
        ElseIf (sfid = 3) Then
            tRow.Controls.Add(CommUtil.CellSet("工單號", 1, 2, False, 0, 0, "center", BColor))
            tRow.Cells.Add(CommUtil.CellSetWithTextBox(1, 2, "txt_machineserialOrwo", 0, 0, 0, NeedInputColor, "center"))

            tRow.Controls.Add(CommUtil.CellSet("出貨日期", 1, 2, False, 0, 0, "center", BColor))
            tRow.Cells.Add(CommUtil.CellSetWithCalenderExtender(1, 2, "installdateOrshipdate", "ce_shipdate", WhiteColor, 0))

            'tCell = New TableCell
            tCell = CommUtil.CellSet("", 1, 4, False, 0, 0, "center", WhiteColor)
            tChk = New CheckBox
            tChk.ID = "chk_firstinstallOrnoassign"
            tChk.Text = "無指定機台"
            tCell.Controls.Add(tChk)
            tRow.Controls.Add(tCell)
        End If
        ContentT.Rows.Add(tRow) 'row=3

        tRow = New TableRow()
        tRow.Font.Size = 16
        tRow.Controls.Add(CommUtil.CellSet("資訊記錄", 1, 16, False, 0, 0, "center", BColor))
        ContentT.Rows.Add(tRow)  'row=4

        '第五列
        tRow = New TableRow()
        tRow.Controls.Add(CommUtil.CellSet("問題類型", 1, 2, False, 0, 0, "center", BColor))
        tRow.Controls.Add(CellSetWithExtender(1, 2, "lb_problemtype", "txt_problemtype", "dde_problemtype", WhiteColor))

        tRow.Controls.Add(CommUtil.CellSet("分類說明", 1, 2, False, 0, 0, "center", BColor))
        tRow.Controls.Add(CellSetWithExtender(1, 2, "lb_typedescrip", "txt_typedescrip", "dde_typedescrip", WhiteColor))

        tRow.Controls.Add(CommUtil.CellSet("版本/規格/序號", 1, 2, False, 0, 0, "center", BColor))
        tRow.Controls.Add(CellSetWithExtender(1, 2, "lb_verandspec", "txt_verandspec", "dde_verandspec", WhiteColor))

        If (sfid = 22) Then
            tRow.Controls.Add(CommUtil.CellSet("當責FAE", 1, 2, False, 0, 0, "center", BColor))
            tRow.Controls.Add(CellSetWithExtender(1, 2, "lb_faeperson", "txt_faeperson", "dde_faeperson", NeedInputColor))
        ElseIf (sfid = 3) Then
            tRow.Controls.Add(CommUtil.CellSet("品管", 1, 2, False, 0, 0, "center", BColor))
            tRow.Controls.Add(CellSetWithExtender(1, 2, "lb_qcperson", "txt_qcperson", "dde_qcperson", NeedInputColor))
        End If
        ContentT.Rows.Add(tRow) 'row=5

        '第六列
        tRow = New TableRow()
        tRow.Controls.Add(CommUtil.CellSet("問題描述", 1, 2, False, 0, 0, "center", BColor))
        tRow.Cells.Add(CommUtil.CellSetWithTextBox(1, 6, "txt_problemdescrip", 5, 0, 400, NeedInputColor, "center"))

        tRow.Controls.Add(CommUtil.CellSet("現場臨時故障排除的處理過程", 1, 2, False, 0, 0, "center", BColor))
        tRow.Cells.Add(CommUtil.CellSetWithTextBox(1, 6, "txt_processdescrip", 5, 0, 400, NeedInputColor, "center"))
        ContentT.Rows.Add(tRow) 'row=6

        '第七列
        tRow = New TableRow()
        tRow.Controls.Add(CommUtil.CellSet("故障品的驗證過程", 1, 2, False, 0, 0, "center", BColor))
        tRow.Cells.Add(CommUtil.CellSetWithTextBox(1, 6, "txt_verifydescrip", 5, 0, 400, WhiteColor, "center"))

        tRow.Controls.Add(CommUtil.CellSet("備註", 1, 2, False, 0, 0, "center", BColor))
        tRow.Cells.Add(CommUtil.CellSetWithTextBox(1, 6, "txt_problemnote", 5, 0, 400, WhiteColor, "center"))
        ContentT.Rows.Add(tRow) 'row=7
    End Sub
    Sub InitTableToSfid3_22ReadOnly()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim BColor, NeedInputColor, WhiteColor As Drawing.Color
        Dim tImage As Image
        Dim tChk As CheckBox
        NeedInputColor = Drawing.Color.AntiqueWhite
        WhiteColor = Drawing.Color.White
        tRow = New TableRow()
        For j = 1 To 16
            tCell = New TableCell
            tCell.BorderWidth = 0
            tCell.Width = 200
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Controls.Add(tCell)
        Next
        FormLogoTitleT.Rows.Add(tRow) 'row=0
        BColor = System.Drawing.Color.LightBlue
        FormLogoTitleT.Font.Name = "標楷體"
        FormLogoTitleT.Font.Size = 12
        tRow = New TableRow()
        tRow.Font.Bold = True
        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Left
        tCell.ColumnSpan = 1
        tImage = New Image
        tImage.ID = "image_logo"
        tImage.ImageUrl = "~/image/jetlog80%.jpg"
        tCell.Controls.Add(tImage)
        tRow.Controls.Add(tCell)
        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.Font.Size = 24
        tCell.ColumnSpan = 14
        If (sfid = 22) Then
            tCell.HorizontalAlign = HorizontalAlign.Center
            tCell.Text = "捷智科技 客戶機台問題反應單"
        ElseIf (sfid = 3) Then
            tCell.HorizontalAlign = HorizontalAlign.Center
            tCell.Text = "捷智科技 廠內機台問題反應單"
        End If
        tRow.Controls.Add(tCell)
        tCell = New TableCell
        tCell.Font.Size = 12
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Right
        tCell.VerticalAlign = VerticalAlign.Bottom
        tCell.ColumnSpan = 1
        If (docnum <> 0) Then
            If (docstatus = "E" Or docstatus = "D") Then
                SqlCmd = "Select convert(varchar(12), docdate, 111) from [dbo].[@XASCH] where docnum=" & docnum
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    tCell.Text = "單號:" & docnum & "<br>建單日期:" & drL(0)
                End If
                drL.Close()
                connL.Close()
            Else
                SqlCmd = "Select convert(varchar(12), signdate, 111) from [dbo].[@XSPWT] where docentry=" & docnum & " and seq=1"
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    tCell.Text = "單號:" & docnum & "<br>送審日期:" & drL(0)
                End If
                drL.Close()
                connL.Close()
            End If
        End If
        tRow.Controls.Add(tCell)
        FormLogoTitleT.Rows.Add(tRow) 'row=1

        tRow = New TableRow()
        For j = 1 To 16
            tCell = New TableCell
            tCell.BorderWidth = 0
            tCell.Width = 200
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Controls.Add(tCell)
        Next
        ContentT.Rows.Add(tRow) 'row=0
        BColor = System.Drawing.Color.LightBlue
        ContentT.Font.Name = "標楷體"
        ContentT.Font.Size = 12
        tRow = New TableRow()
        tRow.Font.Size = 16
        tRow.Controls.Add(CommUtil.CellSet("機台基本資訊", 1, 16, False, 0, 0, "center", BColor))
        ContentT.Rows.Add(tRow)  'row=1
        Dim itemlabelwidth As Integer = 0
        SqlCmd = "Select T0.reportdate,T0.machinetype,T0.cusname,cusfactoryOrmo,model,machineserialOrwo,installdateOrshipdate, " &
             "problemtype,typedescrip,verandspec,faeperson, " &
             "problemdescrip,processdescrip,verifydescrip,problemnote,firstinstallOrnoassign,inwarranty,qcperson " &
             "FROM [dbo].[@XCMRT] T0 WHERE T0.[docentry] =" & docnum
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)

        If (drL.HasRows) Then
            drL.Read()
            '聯絡表頭 
            tRow = New TableRow()
            'CellSet(Text As String, rowspan As Integer, colspan As Integer, FondBold As Boolean, txtid As String, width As Integer, height As Integer, align As String)
            tRow.Controls.Add(CommUtil.CellSet("回報日期", 1, 2, False, itemlabelwidth, 0, "center", BColor))
            tRow.Controls.Add(CommUtil.CellSet(drL(0), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

            tRow.Controls.Add(CommUtil.CellSet("產品別", 1, 2, False, itemlabelwidth, 0, "center", BColor))
            'tRow.Controls.Add(CellSet("", 1, 2, False, 0, 0, "center", False))
            tRow.Controls.Add(CommUtil.CellSet(drL(1), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

            tRow.Controls.Add(CommUtil.CellSet("客戶名稱", 1, 2, False, itemlabelwidth, 0, "center", BColor))
            tRow.Controls.Add(CommUtil.CellSet(drL(2), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

            If (sfid = 22) Then
                tRow.Controls.Add(CommUtil.CellSet("客戶廠區", 1, 2, False, itemlabelwidth, 0, "center", BColor))
                tRow.Controls.Add(CommUtil.CellSet(drL(3), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))
            ElseIf (sfid = 3) Then
                tRow.Controls.Add(CommUtil.CellSet("機台料號", 1, 2, False, itemlabelwidth, 0, "center", BColor))
                tRow.Controls.Add(CommUtil.CellSet(drL(3), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))
            End If
            ContentT.Rows.Add(tRow) 'row=2
            '第三列
            tRow = New TableRow()
            tRow.Controls.Add(CommUtil.CellSet("機台型號", 1, 2, False, itemlabelwidth, 0, "center", BColor))
            tRow.Controls.Add(CommUtil.CellSet(drL(4), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

            If (sfid = 22) Then
                tRow.Controls.Add(CommUtil.CellSet("機台序號", 1, 2, False, itemlabelwidth, 0, "center", BColor))
                tCell = CommUtil.CellSet(drL(5), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor)
                tCell.Font.Size = 10
                tRow.Controls.Add(tCell)

                tRow.Controls.Add(CommUtil.CellSet("裝機日期", 1, 2, False, itemlabelwidth, 0, "center", BColor))
                tRow.Controls.Add(CommUtil.CellSet(drL(6), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

                'tCell = New TableCell
                tCell = CommUtil.CellSet("", 1, 4, False, itemlabelwidth * 2, 0, "center", WhiteColor)
                tChk = New CheckBox
                tChk.ID = "chk_firstinstallOrnoassign"
                tChk.Text = "新安裝&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                If (drL(15) = 1) Then
                    tChk.Checked = True
                Else
                    tChk.Checked = False
                End If
                tCell.Controls.Add(tChk)
                tChk = New CheckBox
                tChk.ID = "chk_inwarranty"
                tChk.Text = "保固期內"
                If (drL(16) = 1) Then
                    tChk.Checked = True
                Else
                    tChk.Checked = False
                End If
                tCell.Controls.Add(tChk)
                tRow.Controls.Add(tCell)
            ElseIf (sfid = 3) Then
                tRow.Controls.Add(CommUtil.CellSet("工單號", 1, 2, False, itemlabelwidth, 0, "center", BColor))
                tCell = CommUtil.CellSet(drL(5), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor)
                tCell.Font.Size = 10
                tRow.Controls.Add(tCell)

                tRow.Controls.Add(CommUtil.CellSet("出貨日期", 1, 2, False, itemlabelwidth, 0, "center", BColor))
                tRow.Controls.Add(CommUtil.CellSet(drL(6), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

                'tCell = New TableCell
                tCell = CommUtil.CellSet("", 1, 4, False, itemlabelwidth * 2, 0, "center", WhiteColor)
                tChk = New CheckBox
                tChk.ID = "chk_firstinstallOrnoassign"
                tChk.Text = "無指定機台"
                If (drL(15) = 1) Then
                    tChk.Checked = True
                Else
                    tChk.Checked = False
                End If
                tCell.Controls.Add(tChk)
                tRow.Controls.Add(tCell)
            End If
            ContentT.Rows.Add(tRow) 'row=3

            tRow = New TableRow()
            tRow.Font.Size = 16
            tRow.Controls.Add(CommUtil.CellSet("資訊記錄", 1, 16, False, 0, 0, "center", BColor))
            ContentT.Rows.Add(tRow)  'row=4

            '第五列
            tRow = New TableRow()
            tRow.Controls.Add(CommUtil.CellSet("問題類型", 1, 2, False, itemlabelwidth, 0, "center", BColor))
            tRow.Controls.Add(CommUtil.CellSet(drL(7), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

            tRow.Controls.Add(CommUtil.CellSet("分類說明", 1, 2, False, itemlabelwidth, 0, "center", BColor))
            tRow.Controls.Add(CommUtil.CellSet(drL(8), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

            tRow.Controls.Add(CommUtil.CellSet("版本/規格/序號", 1, 2, False, itemlabelwidth, 0, "center", BColor))
            tRow.Controls.Add(CommUtil.CellSet(drL(9), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

            If (sfid = 22) Then
                tRow.Controls.Add(CommUtil.CellSet("當責FAE", 1, 2, False, itemlabelwidth, 0, "center", BColor))
                tRow.Controls.Add(CommUtil.CellSet(drL(10), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))
            ElseIf (sfid = 3) Then
                tRow.Controls.Add(CommUtil.CellSet("品管", 1, 2, False, itemlabelwidth, 0, "center", BColor))
                tRow.Controls.Add(CommUtil.CellSet(drL(17), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))
            End If
            ContentT.Rows.Add(tRow) 'row=5

            '第六列
            tRow = New TableRow()
            tRow.Controls.Add(CommUtil.CellSet("問題描述", 1, 2, False, itemlabelwidth, 100, "center", BColor))
            tRow.Controls.Add(CommUtil.CellSet(CommUtil.TextTransToHtmlFormat(drL(11)), 1, 6, False, itemlabelwidth, 100, "left", WhiteColor))

            tRow.Controls.Add(CommUtil.CellSet("現場臨時故障排除的處理過程", 1, 2, False, itemlabelwidth, 100, "center", BColor))
            tRow.Controls.Add(CommUtil.CellSet(CommUtil.TextTransToHtmlFormat(drL(12)), 1, 6, False, itemlabelwidth, 100, "left", WhiteColor))
            ContentT.Rows.Add(tRow) 'row=6

            '第七列
            tRow = New TableRow()
            tRow.Controls.Add(CommUtil.CellSet("故障品的驗證過程", 1, 2, False, itemlabelwidth, 100, "center", BColor))
            tRow.Controls.Add(CommUtil.CellSet(CommUtil.TextTransToHtmlFormat(drL(13)), 1, 6, False, itemlabelwidth, 100, "left", WhiteColor))

            tRow.Controls.Add(CommUtil.CellSet("備註", 1, 2, False, itemlabelwidth, 100, "center", BColor))
            tRow.Controls.Add(CommUtil.CellSet(CommUtil.TextTransToHtmlFormat(drL(14)), 1, 6, False, itemlabelwidth, 100, "left", WhiteColor))
            ContentT.Rows.Add(tRow) 'row=7
        End If
        drL.Close()
        connL.Close()
    End Sub
    'Sub InitTableToSfid3() 'aaaaa
    '    Dim tCell As TableCell
    '    Dim tRow As TableRow
    '    Dim connL As New SqlConnection
    '    Dim drL As SqlDataReader
    '    Dim BColor, NeedInputColor, WhiteColor As Drawing.Color
    '    Dim tImage As Image
    '    Dim tChk As CheckBox
    '    NeedInputColor = Drawing.Color.AntiqueWhite
    '    WhiteColor = Drawing.Color.White
    '    tRow = New TableRow()
    '    For j = 1 To 16
    '        tCell = New TableCell
    '        tCell.BorderWidth = 0
    '        tCell.Width = 200
    '        tCell.HorizontalAlign = HorizontalAlign.Center
    '        tRow.Controls.Add(tCell)
    '    Next
    '    FormLogoTitleT.Rows.Add(tRow) 'row=0
    '    BColor = System.Drawing.Color.LightBlue
    '    FormLogoTitleT.Font.Name = "標楷體"
    '    FormLogoTitleT.Font.Size = 12
    '    tRow = New TableRow()
    '    tRow.Font.Bold = True
    '    tCell = New TableCell
    '    tCell.BorderWidth = 0
    '    tCell.HorizontalAlign = HorizontalAlign.Left
    '    tCell.ColumnSpan = 1
    '    tImage = New Image
    '    tImage.ID = "image_logo"
    '    tImage.ImageUrl = "~/image/jetlog80%.jpg"
    '    tCell.Controls.Add(tImage)
    '    tRow.Controls.Add(tCell)
    '    tCell = New TableCell
    '    tCell.BorderWidth = 0
    '    tCell.Font.Size = 24
    '    tCell.ColumnSpan = 14
    '    tCell.HorizontalAlign = HorizontalAlign.Center
    '    tCell.Text = "捷智科技 廠內機台問題反應單"
    '    tRow.Controls.Add(tCell)
    '    tCell = New TableCell
    '    tCell.Font.Size = 12
    '    tCell.BorderWidth = 0
    '    tCell.HorizontalAlign = HorizontalAlign.Right
    '    tCell.VerticalAlign = VerticalAlign.Bottom
    '    tCell.ColumnSpan = 1
    '    If (docnum <> 0) Then
    '        If (docstatus = "E" Or docstatus = "D") Then
    '            SqlCmd = "Select convert(varchar(12), docdate, 111) from [dbo].[@XASCH] where docnum=" & docnum
    '            drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
    '            If (drL.HasRows) Then
    '                drL.Read()
    '                tCell.Text = "單號:" & docnum & "<br>建單日期:" & drL(0)
    '            End If
    '            drL.Close()
    '            connL.Close()
    '        Else
    '            SqlCmd = "Select convert(varchar(12), signdate, 111) from [dbo].[@XSPWT] where docentry=" & docnum & " and seq=1"
    '            drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
    '            If (drL.HasRows) Then
    '                drL.Read()
    '                tCell.Text = "單號:" & docnum & "<br>送審日期:" & drL(0)
    '            End If
    '            drL.Close()
    '            connL.Close()
    '        End If
    '    End If
    '    tRow.Controls.Add(tCell)
    '    FormLogoTitleT.Rows.Add(tRow) 'row=1

    '    tRow = New TableRow()
    '    For j = 1 To 16
    '        tCell = New TableCell
    '        tCell.BorderWidth = 0
    '        tCell.Width = 200
    '        tCell.HorizontalAlign = HorizontalAlign.Center
    '        tRow.Controls.Add(tCell)
    '    Next
    '    ContentT.Rows.Add(tRow) 'row=0
    '    BColor = System.Drawing.Color.LightBlue
    '    ContentT.Font.Name = "標楷體"
    '    ContentT.Font.Size = 14
    '    tRow = New TableRow()
    '    tRow.Controls.Add(CommUtil.CellSet("機台基本資訊", 1, 16, False, 0, 0, "center", BColor))
    '    ContentT.Rows.Add(tRow)  'row=1
    '    '聯絡表頭 
    '    tRow = New TableRow()
    '    'CellSet(Text As String, rowspan As Integer, colspan As Integer, FondBold As Boolean, txtid As String, width As Integer, height As Integer, align As String)
    '    tRow.Controls.Add(CommUtil.CellSet("回報日期", 1, 2, False, 0, 0, "center", BColor))
    '    tRow.Cells.Add(CommUtil.CellSetWithCalenderExtender(1, 2, "txt_reportdate", "ce_reportdate", NeedInputColor, 0))

    '    tRow.Controls.Add(CommUtil.CellSet("產品別", 1, 2, False, 0, 0, "center", BColor))
    '    'tRow.Controls.Add(CellSet("", 1, 2, False, 0, 0, "center", False))
    '    tRow.Controls.Add(CellSetWithExtender(1, 2, "lb_machinetype", "txt_machinetype", "dde_machinetype", NeedInputColor))

    '    tRow.Controls.Add(CommUtil.CellSet("客戶名稱", 1, 2, False, 0, 0, "center", BColor))
    '    tRow.Controls.Add(CellSetWithExtender(1, 2, "lb_cusname", "txt_cusname", "dde_cusname", NeedInputColor))

    '    tRow.Controls.Add(CommUtil.CellSet("機台料號", 1, 2, False, 0, 0, "center", BColor))
    '    tRow.Cells.Add(CommUtil.CellSetWithTextBox(1, 2, "txt_mo", 0, 0, 0, NeedInputColor, "center"))
    '    ContentT.Rows.Add(tRow) 'row=2
    '    '第三列
    '    tRow = New TableRow()
    '    tRow.Controls.Add(CommUtil.CellSet("機台型號", 1, 2, False, 0, 0, "center", BColor))
    '    tRow.Controls.Add(CellSetWithExtender(1, 2, "lb_model", "txt_model", "dde_model", NeedInputColor))

    '    tRow.Controls.Add(CommUtil.CellSet("工單號", 1, 2, False, 0, 0, "center", BColor))
    '    tRow.Cells.Add(CommUtil.CellSetWithTextBox(1, 2, "txt_wo", 0, 0, 0, NeedInputColor, "center"))

    '    tRow.Controls.Add(CommUtil.CellSet("出貨日期", 1, 2, False, 0, 0, "center", BColor))
    '    tRow.Cells.Add(CommUtil.CellSetWithCalenderExtender(1, 2, "txt_shipdate", "ce_shipdate", WhiteColor, 0))

    '    'tCell = New TableCell
    '    tCell = CommUtil.CellSet("", 1, 4, False, 0, 0, "center", WhiteColor)
    '    tChk = New CheckBox
    '    tChk.ID = "chk_noassign"
    '    tChk.Text = "無指定機台"
    '    tCell.Controls.Add(tChk)
    '    tRow.Controls.Add(tCell)
    '    ContentT.Rows.Add(tRow) 'row=3

    '    tRow = New TableRow()
    '    tRow.Controls.Add(CommUtil.CellSet("資訊記錄", 1, 16, False, 0, 0, "center", BColor))
    '    ContentT.Rows.Add(tRow)  'row=4

    '    '第五列
    '    tRow = New TableRow()
    '    tRow.Controls.Add(CommUtil.CellSet("問題類型", 1, 2, False, 0, 0, "center", BColor))
    '    tRow.Controls.Add(CellSetWithExtender(1, 2, "lb_problemtype", "txt_problemtype", "dde_problemtype", WhiteColor))

    '    tRow.Controls.Add(CommUtil.CellSet("分類說明", 1, 2, False, 0, 0, "center", BColor))
    '    tRow.Controls.Add(CellSetWithExtender(1, 2, "lb_typedescrip", "txt_typedescrip", "dde_typedescrip", WhiteColor))

    '    tRow.Controls.Add(CommUtil.CellSet("版本/規格/序號", 1, 2, False, 0, 0, "center", BColor))
    '    tRow.Controls.Add(CellSetWithExtender(1, 2, "lb_verandspec", "txt_verandspec", "dde_verandspec", WhiteColor))

    '    tRow.Controls.Add(CommUtil.CellSet("品管", 1, 2, False, 0, 0, "center", BColor))
    '    tRow.Controls.Add(CellSetWithExtender(1, 2, "lb_qcperson", "txt_qcperson", "dde_qcperson", NeedInputColor))
    '    ContentT.Rows.Add(tRow) 'row=5

    '    '第六列
    '    tRow = New TableRow()
    '    tRow.Controls.Add(CommUtil.CellSet("問題描述", 1, 2, False, 0, 0, "center", BColor))
    '    tRow.Cells.Add(CommUtil.CellSetWithTextBox(1, 6, "txt_problemdescrip", 5, 0, 400, NeedInputColor, "center"))

    '    tRow.Controls.Add(CommUtil.CellSet("現場臨時故障排除的處理過程", 1, 2, False, 0, 0, "center", BColor))
    '    tRow.Cells.Add(CommUtil.CellSetWithTextBox(1, 6, "txt_processdescrip", 5, 0, 400, NeedInputColor, "center"))
    '    ContentT.Rows.Add(tRow) 'row=6

    '    '第七列
    '    tRow = New TableRow()
    '    tRow.Controls.Add(CommUtil.CellSet("故障品的驗證過程", 1, 2, False, 0, 0, "center", BColor))
    '    tRow.Cells.Add(CommUtil.CellSetWithTextBox(1, 6, "txt_verifydescrip", 5, 0, 400, WhiteColor, "center"))

    '    tRow.Controls.Add(CommUtil.CellSet("備註", 1, 2, False, 0, 0, "center", BColor))
    '    tRow.Cells.Add(CommUtil.CellSetWithTextBox(1, 6, "txt_problemnote", 5, 0, 400, WhiteColor, "center"))
    '    ContentT.Rows.Add(tRow) 'row=7
    'End Sub
    'Sub InitTableToSfid3ReadOnly()
    '    Dim tCell As TableCell
    '    Dim tRow As TableRow
    '    Dim connL As New SqlConnection
    '    Dim drL As SqlDataReader
    '    Dim BColor, NeedInputColor, WhiteColor As Drawing.Color
    '    Dim tImage As Image
    '    Dim tChk As CheckBox
    '    NeedInputColor = Drawing.Color.AntiqueWhite
    '    WhiteColor = Drawing.Color.White
    '    tRow = New TableRow()
    '    For j = 1 To 16
    '        tCell = New TableCell
    '        tCell.BorderWidth = 0
    '        tCell.Width = 200
    '        tCell.HorizontalAlign = HorizontalAlign.Center
    '        tRow.Controls.Add(tCell)
    '    Next
    '    FormLogoTitleT.Rows.Add(tRow) 'row=0
    '    BColor = System.Drawing.Color.LightBlue
    '    FormLogoTitleT.Font.Name = "標楷體"
    '    FormLogoTitleT.Font.Size = 12
    '    tRow = New TableRow()
    '    tRow.Font.Bold = True
    '    tCell = New TableCell
    '    tCell.BorderWidth = 0
    '    tCell.HorizontalAlign = HorizontalAlign.Left
    '    tCell.ColumnSpan = 1
    '    tImage = New Image
    '    tImage.ID = "image_logo"
    '    tImage.ImageUrl = "~/image/jetlog80%.jpg"
    '    tCell.Controls.Add(tImage)
    '    tRow.Controls.Add(tCell)
    '    tCell = New TableCell
    '    tCell.BorderWidth = 0
    '    tCell.Font.Size = 24
    '    tCell.ColumnSpan = 14
    '    tCell.HorizontalAlign = HorizontalAlign.Center
    '    tCell.Text = "捷智科技 廠內機台問題反應單"
    '    tRow.Controls.Add(tCell)
    '    tCell = New TableCell
    '    tCell.Font.Size = 12
    '    tCell.BorderWidth = 0
    '    tCell.HorizontalAlign = HorizontalAlign.Right
    '    tCell.VerticalAlign = VerticalAlign.Bottom
    '    tCell.ColumnSpan = 1
    '    If (docnum <> 0) Then
    '        If (docstatus = "E" Or docstatus = "D") Then
    '            SqlCmd = "Select convert(varchar(12), docdate, 111) from [dbo].[@XASCH] where docnum=" & docnum
    '            drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
    '            If (drL.HasRows) Then
    '                drL.Read()
    '                tCell.Text = "單號:" & docnum & "<br>建單日期:" & drL(0)
    '            End If
    '            drL.Close()
    '            connL.Close()
    '        Else
    '            SqlCmd = "Select convert(varchar(12), signdate, 111) from [dbo].[@XSPWT] where docentry=" & docnum & " and seq=1"
    '            drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
    '            If (drL.HasRows) Then
    '                drL.Read()
    '                tCell.Text = "單號:" & docnum & "<br>送審日期:" & drL(0)
    '            End If
    '            drL.Close()
    '            connL.Close()
    '        End If
    '    End If
    '    tRow.Controls.Add(tCell)
    '    FormLogoTitleT.Rows.Add(tRow) 'row=1

    '    tRow = New TableRow()
    '    For j = 1 To 16
    '        tCell = New TableCell
    '        tCell.BorderWidth = 0
    '        tCell.Width = 200
    '        tCell.HorizontalAlign = HorizontalAlign.Center
    '        tRow.Controls.Add(tCell)
    '    Next
    '    ContentT.Rows.Add(tRow) 'row=0
    '    BColor = System.Drawing.Color.LightBlue
    '    ContentT.Font.Name = "標楷體"
    '    ContentT.Font.Size = 12
    '    tRow = New TableRow()
    '    tRow.Controls.Add(CommUtil.CellSet("機台基本資訊", 1, 16, False, 0, 0, "center", BColor))
    '    ContentT.Rows.Add(tRow)  'row=1
    '    Dim itemlabelwidth As Integer = 0
    '    SqlCmd = "Select T0.reportdate,T0.machinetype,T0.cusname,mo,model,wo,shipdate, " &
    '         "problemtype,typedescrip,verandspec,qcperson, " &
    '         "problemdescrip,processdescrip,verifydescrip,problemnote,noassign " &
    '         "FROM [dbo].[@XFMRT] T0 WHERE T0.[docentry] =" & docnum
    '    drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)

    '    If (drL.HasRows) Then
    '        drL.Read()
    '        '聯絡表頭 
    '        tRow = New TableRow()
    '        'CellSet(Text As String, rowspan As Integer, colspan As Integer, FondBold As Boolean, txtid As String, width As Integer, height As Integer, align As String)
    '        tRow.Controls.Add(CommUtil.CellSet("回報日期", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        tRow.Controls.Add(CommUtil.CellSet(drL(0), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

    '        tRow.Controls.Add(CommUtil.CellSet("產品別", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        'tRow.Controls.Add(CellSet("", 1, 2, False, 0, 0, "center", False))
    '        tRow.Controls.Add(CommUtil.CellSet(drL(1), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

    '        tRow.Controls.Add(CommUtil.CellSet("客戶名稱", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        tRow.Controls.Add(CommUtil.CellSet(drL(2), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

    '        tRow.Controls.Add(CommUtil.CellSet("機台料號", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        tRow.Controls.Add(CommUtil.CellSet(drL(3), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))
    '        ContentT.Rows.Add(tRow) 'row=2
    '        '第三列
    '        tRow = New TableRow()
    '        tRow.Controls.Add(CommUtil.CellSet("機台型號", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        tRow.Controls.Add(CommUtil.CellSet(drL(4), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

    '        tRow.Controls.Add(CommUtil.CellSet("工單號", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        tCell = CommUtil.CellSet(drL(5), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor)
    '        tCell.Font.Size = 10
    '        tRow.Controls.Add(tCell)

    '        tRow.Controls.Add(CommUtil.CellSet("出貨日期", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        tRow.Controls.Add(CommUtil.CellSet(drL(6), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

    '        'tCell = New TableCell
    '        tCell = CommUtil.CellSet("", 1, 4, False, itemlabelwidth * 2, 0, "center", WhiteColor)
    '        tChk = New CheckBox
    '        tChk.ID = "chk_firstinstallOrnoassign"
    '        tChk.Text = "無指定機台"
    '        If (drL(15) = 1) Then
    '            tChk.Checked = True
    '        Else
    '            tChk.Checked = False
    '        End If
    '        tCell.Controls.Add(tChk)
    '        tRow.Controls.Add(tCell)

    '        ContentT.Rows.Add(tRow) 'row=3

    '        tRow = New TableRow()
    '        tRow.Controls.Add(CommUtil.CellSet("資訊記錄", 1, 16, False, 0, 0, "center", BColor))
    '        ContentT.Rows.Add(tRow)  'row=4

    '        '第五列
    '        tRow = New TableRow()
    '        tRow.Controls.Add(CommUtil.CellSet("問題類型", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        tRow.Controls.Add(CommUtil.CellSet(drL(7), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

    '        tRow.Controls.Add(CommUtil.CellSet("分類說明", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        tRow.Controls.Add(CommUtil.CellSet(drL(8), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

    '        tRow.Controls.Add(CommUtil.CellSet("版本/規格/序號", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        tRow.Controls.Add(CommUtil.CellSet(drL(9), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

    '        tRow.Controls.Add(CommUtil.CellSet("品管", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        tRow.Controls.Add(CommUtil.CellSet(drL(10), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))
    '        ContentT.Rows.Add(tRow) 'row=5

    '        '第六列
    '        tRow = New TableRow()
    '        tRow.Controls.Add(CommUtil.CellSet("問題描述", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        tRow.Controls.Add(CommUtil.CellSet(drL(11), 1, 6, False, itemlabelwidth, 0, "left", WhiteColor))

    '        tRow.Controls.Add(CommUtil.CellSet("現場臨時故障排除的處理過程", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        tRow.Controls.Add(CommUtil.CellSet(drL(12), 1, 6, False, itemlabelwidth, 0, "left", WhiteColor))
    '        ContentT.Rows.Add(tRow) 'row=6

    '        '第七列
    '        tRow = New TableRow()
    '        tRow.Controls.Add(CommUtil.CellSet("故障品的驗證過程", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        tRow.Controls.Add(CommUtil.CellSet(drL(13), 1, 6, False, itemlabelwidth, 0, "left", WhiteColor))

    '        tRow.Controls.Add(CommUtil.CellSet("備註", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        tRow.Controls.Add(CommUtil.CellSet(drL(14), 1, 6, False, itemlabelwidth, 0, "left", WhiteColor))
    '        ContentT.Rows.Add(tRow) 'row=7
    '    End If
    '    drL.Close()
    '    connL.Close()
    'End Sub
    Sub InitTableToSfid16()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim tTxt As TextBox
        Dim Labelx As Label
        Dim dDDL As DropDownList
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim BColor As Drawing.Color
        Dim tImage As Image
        'BColor = System.Drawing.Color.LightBlue
        'ContentT.Font.Name = "標楷體"
        'ContentT.Font.Size = 16
        'tRow = New TableRow()
        'tRow.Font.Bold = True
        'tCell = New TableCell
        'tCell.BorderWidth = 0
        'tCell.ColumnSpan = 6
        'tCell.HorizontalAlign = HorizontalAlign.Center
        'tCell.Text = "Jet門禁磁卡補刷卡單"
        'tRow.Controls.Add(tCell)
        'ContentT.Rows.Add(tRow)

        tRow = New TableRow()
        For j = 1 To 10
            tCell = New TableCell
            tCell.BorderWidth = 0
            tCell.Width = 200
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Controls.Add(tCell)
        Next
        FormLogoTitleT.Rows.Add(tRow) 'row=0
        BColor = System.Drawing.Color.LightBlue
        FormLogoTitleT.Font.Name = "標楷體"
        FormLogoTitleT.Font.Size = 12
        tRow = New TableRow()
        tRow.Font.Bold = True
        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Left
        tCell.ColumnSpan = 1
        tImage = New Image
        tImage.ID = "image_logo"
        tImage.ImageUrl = "~/image/jetlog80%.jpg"
        tCell.Controls.Add(tImage)
        tRow.Controls.Add(tCell)
        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.Font.Size = 24
        tCell.ColumnSpan = 8
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "捷智科技 門禁磁卡補刷卡單"
        tRow.Controls.Add(tCell)
        tCell = New TableCell
        tCell.Font.Size = 12
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Right
        tCell.VerticalAlign = VerticalAlign.Bottom
        tCell.ColumnSpan = 1
        If (docnum <> 0) Then
            If (docstatus = "E" Or docstatus = "D") Then
                SqlCmd = "Select convert(varchar(12), docdate, 111) from [dbo].[@XASCH] where docnum=" & docnum
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    tCell.Text = "單號:" & docnum & "<br>建單日期:" & drL(0)
                End If
                drL.Close()
                connL.Close()
            Else
                SqlCmd = "Select convert(varchar(12), signdate, 111) from [dbo].[@XSPWT] where docentry=" & docnum & " and seq=1"
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    tCell.Text = "單號:" & docnum & "<br>送審日期:" & drL(0)
                End If
                drL.Close()
                connL.Close()
            End If
        End If
        tRow.Controls.Add(tCell)
        FormLogoTitleT.Rows.Add(tRow) 'row=1

        For i = 1 To 6
            tRow = New TableRow()
            'tRow.BackColor = Drawing.Color.LightGreen
            tCell = New TableCell
            tCell.BorderWidth = 1
            'tCell.Width = 120
            tCell.HorizontalAlign = HorizontalAlign.Left
            If (i = 1) Then
                tCell.Text = "填寫日期"
            ElseIf (i = 2) Then
                tCell.Text = "補刷卡日期"
            ElseIf (i = 3) Then
                tCell.Text = "補刷卡時間"
            ElseIf (i = 4) Then
                tCell.Text = "補刷卡事由"
            ElseIf (i = 5) Then
                tCell.Text = "補刷卡人"
                tCell.ColumnSpan = 3
                tCell.HorizontalAlign = HorizontalAlign.Center
            ElseIf (i = 6) Then
                tCell.ColumnSpan = 3
                tCell.HorizontalAlign = HorizontalAlign.Center
                If (docstatus = "A" Or docstatus = "E" Or docstatus = "D" Or docstatus = "B" Or docstatus = "R") Then
                    'MsgBox(docstatus)
                    SqlCmd = "select id ,name from dbo.[user] where denyf=0 and branch='" & Session("branch") & "' order by grp"
                    dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn) 'modify
                    dDDL = New DropDownList()
                    dDDL.ID = "ddlrsuser_" & i & "_0"
                    dDDL.Items.Clear()
                    dDDL.Items.Add("請選擇")
                    If (dr.HasRows) Then
                        Do While (dr.Read())
                            dDDL.Items.Add(dr(0) & " " & dr(1))
                        Loop
                    End If
                    If (Session("branch") <> "CNC") Then
                        dDDL.SelectedValue = Session("s_id") & " " & Session("s_name")
                    End If
                    dr.Close()
                    conn.Close()
                    dDDL.BackColor = BColor
                    tCell.Controls.Add(dDDL)
                End If
            End If
            tRow.Controls.Add(tCell)

            tCell = New TableCell
            tCell.BorderWidth = 1
            'tCell.Width = 120
            tCell.HorizontalAlign = HorizontalAlign.Left
            If (i = 5) Then
                tCell.Text = "正航代號"
                tCell.ColumnSpan = 3
                tCell.HorizontalAlign = HorizontalAlign.Center
            ElseIf (i = 6) Then
                tCell.ColumnSpan = 3
                tCell.HorizontalAlign = HorizontalAlign.Center
                If (docstatus = "A" Or docstatus = "E" Or docstatus = "D" Or docstatus = "B" Or docstatus = "R") Then
                    tTxt = New TextBox
                    tTxt.ID = "txtv5usercode_" & i & "_1"
                    tTxt.Width = 120
                    tTxt.BackColor = BColor
                    tCell.Controls.Add(tTxt)
                End If
            ElseIf (i = 1 Or i = 2 Or i = 3 Or i = 4) Then
                tCell.ColumnSpan = 5
                If (docstatus = "A" Or docstatus = "E" Or docstatus = "D" Or docstatus = "B" Or docstatus = "R") Then
                    If (i = 1 And docstatus = "A") Then
                        tCell.Text = "&nbsp&nbsp" & Now.Year & " 年 " & Now.Month & " 月 " & Now.Day & " 日"
                    ElseIf (i = 2) Then
                        'MsgBox(docstatus)
                        dDDL = New DropDownList()
                        dDDL.ID = "ddlfromyear_" & i & "_1"
                        dDDL.Items.Clear()
                        dDDL.Items.Add(Now.Year - 1)
                        dDDL.Items.Add(Now.Year)
                        dDDL.SelectedValue = Now.Year
                        dDDL.BackColor = BColor
                        tCell.Controls.Add(dDDL)
                        Labelx = New Label
                        Labelx.ID = "labelfromyear_" & i & "_1"
                        Labelx.Text = "&nbsp&nbsp年&nbsp&nbsp"
                        tCell.Controls.Add(Labelx)

                        dDDL = New DropDownList()
                        dDDL.ID = "ddlfrommonth_" & i & "_1"
                        dDDL.Items.Clear()
                        For k = 1 To 12
                            dDDL.Items.Add(k.ToString.PadLeft(2, "0"))
                        Next
                        If (CInt(Now.Month) >= 10) Then
                            dDDL.SelectedValue = Now.Month
                        Else
                            dDDL.SelectedValue = "0" & Now.Month
                        End If
                        dDDL.BackColor = BColor
                        tCell.Controls.Add(dDDL)
                        Labelx = New Label
                        Labelx.ID = "labelfrommonth_" & i & "_1"
                        Labelx.Text = "&nbsp&nbsp月&nbsp&nbsp"
                        tCell.Controls.Add(Labelx)

                        dDDL = New DropDownList()
                        dDDL.ID = "ddlfromday_" & i & "_1"
                        dDDL.Items.Clear()
                        dDDL.Items.Add("選擇")
                        For k = 1 To 31
                            dDDL.Items.Add(k.ToString.PadLeft(2, "0"))
                        Next
                        dDDL.BackColor = BColor
                        tCell.Controls.Add(dDDL)
                        Labelx = New Label
                        Labelx.ID = "labelfromday_" & i & "_1"
                        Labelx.Text = "&nbsp&nbsp日&nbsp&nbsp&nbsp&nbsp至&nbsp&nbsp&nbsp&nbsp"
                        tCell.Controls.Add(Labelx)

                        dDDL = New DropDownList()
                        dDDL.ID = "ddltoyear_" & i & "_1"
                        dDDL.Items.Clear()
                        dDDL.Items.Add(Now.Year - 1)
                        dDDL.Items.Add(Now.Year)
                        dDDL.SelectedValue = Now.Year
                        dDDL.BackColor = BColor
                        tCell.Controls.Add(dDDL)
                        Labelx = New Label
                        Labelx.ID = "labeltoyear_" & i & "_1"
                        Labelx.Text = "&nbsp&nbsp年&nbsp&nbsp"
                        tCell.Controls.Add(Labelx)

                        dDDL = New DropDownList()
                        dDDL.ID = "ddltomonth_" & i & "_1"
                        dDDL.Items.Clear()
                        For k = 1 To 12
                            dDDL.Items.Add(k.ToString.PadLeft(2, "0"))
                        Next
                        If (CInt(Now.Month) >= 10) Then
                            dDDL.SelectedValue = Now.Month
                        Else
                            dDDL.SelectedValue = "0" & Now.Month
                        End If
                        dDDL.BackColor = BColor
                        tCell.Controls.Add(dDDL)
                        Labelx = New Label
                        Labelx.ID = "labeltomonth_" & i & "_1"
                        Labelx.Text = "&nbsp&nbsp月&nbsp&nbsp"
                        tCell.Controls.Add(Labelx)

                        dDDL = New DropDownList()
                        dDDL.ID = "ddltoday_" & i & "_1"
                        dDDL.Items.Clear()
                        dDDL.Items.Add("選擇")
                        For k = 1 To 31
                            dDDL.Items.Add(k.ToString.PadLeft(2, "0"))
                        Next
                        dDDL.BackColor = BColor
                        tCell.Controls.Add(dDDL)
                        Labelx = New Label
                        Labelx.ID = "labeltoday_" & i & "_1"
                        Labelx.Text = "&nbsp&nbsp日&nbsp&nbsp止"
                        tCell.Controls.Add(Labelx)
                    ElseIf (i = 3) Then
                        Labelx = New Label
                        Labelx.ID = "labelbegin_" & i & "_1"
                        Labelx.Text = "&nbsp&nbsp起始&nbsp&nbsp"
                        tCell.Controls.Add(Labelx)

                        dDDL = New DropDownList()
                        dDDL.ID = "ddlfromhour_" & i & "_1"
                        dDDL.Items.Clear()
                        For k = 1 To 24
                            dDDL.Items.Add(k.ToString.PadLeft(2, "0"))
                        Next
                        dDDL.SelectedValue = "08"
                        dDDL.BackColor = BColor
                        tCell.Controls.Add(dDDL)
                        Labelx = New Label
                        Labelx.ID = "labelfromhour_" & i & "_1"
                        Labelx.Text = "&nbsp&nbsp時&nbsp&nbsp"
                        tCell.Controls.Add(Labelx)

                        dDDL = New DropDownList()
                        dDDL.ID = "ddlfrommin_" & i & "_1"
                        dDDL.Items.Clear()
                        dDDL.Items.Add("00")
                        dDDL.Items.Add("30")
                        dDDL.SelectedValue = "30"
                        dDDL.BackColor = BColor
                        tCell.Controls.Add(dDDL)
                        Labelx = New Label
                        Labelx.ID = "labelfrommin_" & i & "_1"
                        Labelx.Text = "&nbsp&nbsp分&nbsp&nbsp&nbsp&nbsp至&nbsp&nbsp&nbsp&nbsp"
                        tCell.Controls.Add(Labelx)

                        dDDL = New DropDownList()
                        dDDL.ID = "ddltohour_" & i & "_1"
                        dDDL.Items.Clear()
                        For k = 1 To 24
                            dDDL.Items.Add(k.ToString.PadLeft(2, "0"))
                        Next
                        dDDL.SelectedValue = "17"
                        dDDL.BackColor = BColor
                        tCell.Controls.Add(dDDL)
                        Labelx = New Label
                        Labelx.ID = "labeltohour_" & i & "_1"
                        Labelx.Text = "&nbsp&nbsp時&nbsp&nbsp"
                        tCell.Controls.Add(Labelx)

                        dDDL = New DropDownList()
                        dDDL.ID = "ddltomin_" & i & "_1"
                        dDDL.Items.Clear()
                        dDDL.Items.Add("00")
                        dDDL.Items.Add("30")
                        dDDL.SelectedValue = "30"
                        dDDL.BackColor = BColor
                        tCell.Controls.Add(dDDL)
                        Labelx = New Label
                        Labelx.ID = "labeltomin_" & i & "_1"
                        Labelx.Text = "&nbsp&nbsp分&nbsp&nbsp&nbsp&nbsp止&nbsp&nbsp&nbsp&nbsp"
                        tCell.Controls.Add(Labelx)
                    ElseIf (i = 4) Then
                        tTxt = New TextBox
                        tTxt.ID = "txtrsreason_" & i & "_1"
                        tTxt.Width = 700
                        tTxt.BackColor = BColor
                        tCell.Controls.Add(tTxt)
                    End If
                End If
            End If
            tRow.Controls.Add(tCell)
            ContentT.Rows.Add(tRow)
        Next
        tRow = New TableRow()
        tRow.Font.Bold = True
        For j = 1 To 6
            tCell = New TableCell
            tCell.BorderWidth = 0
            tCell.Width = 100
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Controls.Add(tCell)
        Next
        ContentT.Rows.Add(tRow)
    End Sub

    Sub ContentTCreate()
        If (sfid = 51 Or sfid = 50 Or sfid = 49 Or sfid = 100) Then
            'MsgBox("1-" & Request.QueryString("num"))
            InitTableToSfid49_50_51_100()
            ShowMaterialList(Request.QueryString("num"))
            'MsgBox("2")
        ElseIf (sfid = 16) Then '未刷卡單
            InitTableToSfid16()
            ShowXRSCT()
        ElseIf (sfid = 1) Then '通用聯絡單
            InitTableToSfid1()
            ShowXGCT()
        ElseIf (sfid = 12) Then '機台聯絡單單
            InitTableToSfid12()
            ShowXMSCT()
        ElseIf (sfid = 22 Or sfid = 3) Then '客戶(廠內)機台問題反應單
            If (sid_create = sid) Then
                InitTableToSfid3_22()
                WriteListBoxItemForXCMRT()
                ShowXCMRT()
            Else
                InitTableToSfid3_22ReadOnly()
            End If
        ElseIf (sfid = 23 Or sfid = 24) Then '離倉料件管制單
            ChkUsingAttach.Visible = False
            InitTableToSfid23_24()
            ShowMaterialList23_24(Request.QueryString("num"))
        ElseIf (sfid = 101) Then '料件還回單
            InitTableToSfid101() '資料導入已做在其中
            'If (TxtAttaDoc.Text <> "NA" And TxtAttaDoc.Text <> "") Then
            '    ShowMaterialList101(CLng(TxtAttaDoc.Text))
            'End If
        End If
    End Sub
    Sub ShowMaterialList(modifynum As Long) 'ron
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim totalprice As Double
        Dim descrip As String = ""
        '        Dim HyperBtn As LinkButton
        '        Dim ChkMaterialDel As CheckBox
        Dim i As Integer
        SqlCmd = "Select descrip FROM [dbo].[@XSMLS] T0 WHERE T0.[docentry] =" & docnum & " and head=1"
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            drL.Read()
            descrip = drL(0)
        End If
        drL.Close()
        connL.Close()
        TxtReason.Text = descrip
        If ((System.Text.RegularExpressions.Regex.Matches(descrip, "\r\n").Count + 1) > 7) Then
            TxtReason.Rows = System.Text.RegularExpressions.Regex.Matches(descrip, "\r\n").Count + 1
        End If
        SqlCmd = "Select IsNull(sum(quantity*price),0) FROM [dbo].[@XSMLS] T0 WHERE head=0 and T0.[docentry] =" & docnum
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        drL.Read()
        totalprice = drL(0)
        drL.Close()
        connL.Close()
        i = 4 '料件再Table 之起始列
        SqlCmd = "Select T0.itemcode,T0.itemname,T0.quantity,T0.price,T0.method,T0.comment,T0.num, " &
                "T2.Onhand,T2.IsCommited,T2.OnOrder,T1.OnHand-T2.OnHand,T1.IsCommited-T2.IsCommited,T1.OnOrder-T2.OnOrder " &
                "FROM [dbo].[@XSMLS] T0 " &
                "Inner Join OITM T1 On T0.itemcode=T1.Itemcode Inner Join OITW T2 on T1.Itemcode=T2.Itemcode " &
                "WHERE T0.head=0 And T0.[docentry] =" & docnum & " And T2.Whscode='" & Session("usingwhs") & "' ORDER BY T0.num"
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            Do While (drL.Read())
                ContentT.Rows(i).Cells(0).Text = i - 3
                ContentT.Rows(i).Cells(1).Text = drL(0)
                ContentT.Rows(i).Cells(2).Text = drL(1)
                ContentT.Rows(i).Cells(3).Text = drL(2)
                ContentT.Rows(i).Cells(4).Text = Format(drL(3), "###,###.##")
                ContentT.Rows(i).Cells(5).Text = Format(drL(3) * drL(2), "###,###.##") '總價
                ContentT.Rows(i).Cells(6).Text = CInt(drL(7)) '庫存
                ContentT.Rows(i).Cells(7).Text = CInt(drL(8)) '本倉需求
                ContentT.Rows(i).Cells(8).Text = CInt(drL(9)) '本倉供給
                ContentT.Rows(i).Cells(9).Text = CInt(drL(10)) '它倉庫存
                ContentT.Rows(i).Cells(10).Text = CInt(drL(11)) '它廠需求
                ContentT.Rows(i).Cells(11).Text = CInt(drL(12)) '它廠供給
                ContentT.Rows(i).Cells(12).Text = drL(4) '報廢原因
                ContentT.Rows(i).Cells(13).Text = drL(5) '備註
                If (drL(2) > drL(7) + drL(9)) Then
                    ContentT.Rows(i).Cells(3).BackColor = Drawing.Color.Red
                ElseIf (drL(2) > drL(7)) Then
                    ContentT.Rows(i).Cells(3).BackColor = Drawing.Color.Yellow
                End If
                If (docstatus = "B" Or docstatus = "E" Or docstatus = "D" Or docstatus = "R" Or docstatus = "A") Then
                    If (CType(ContentT.FindControl("chkdel_" & i & "_9"), CheckBox).Checked) Then
                        CType(ContentT.FindControl("hypermodify_" & i & "_8"), HyperLink).NavigateUrl = "cLsignoff.aspx?smid=sg&smode=2&act=material_del&status=" & docstatus &
                    "&docnum=" & docnum & "&formstatusindex=" & formstatusindex &
                    "&formtypeindex=" & formtypeindex & "&sfid=" & sfid & "&subject=" & TxtSubject.Text & "&num=" & drL(6)
                        CType(ContentT.FindControl("hypermodify_" & i & "_8"), HyperLink).Text = "刪除"
                    Else
                        If (drL(6) <> modifynum) Then
                            CType(ContentT.FindControl("hypermodify_" & i & "_8"), HyperLink).NavigateUrl = "cLsignoff.aspx?smid=sg&smode=2&act=material_modify&status=" & docstatus &
                        "&docnum=" & docnum & "&formstatusindex=" & formstatusindex &
                        "&formtypeindex=" & formtypeindex & "&sfid=" & sfid & "&subject=" & TxtSubject.Text & "&num=" & drL(6)
                            CType(ContentT.FindControl("hypermodify_" & i & "_8"), HyperLink).Text = "修改"
                            CType(ContentT.FindControl("hypermodify_" & i & "_8"), HyperLink).Enabled = True
                        Else
                            CType(ContentT.FindControl("hypermodify_" & i & "_8"), HyperLink).Text = "修改中.."
                            CType(ContentT.FindControl("hypermodify_" & i & "_8"), HyperLink).Enabled = False
                        End If
                    End If
                End If
                i = i + 1
            Loop
            ContentT.Rows(i).Cells(4).Text = "Total"
            ContentT.Rows(i).Cells(5).Text = Format(totalprice, "###,###.##")
            TxtPrice.Text = Format(totalprice, "###,###.##")
        Else
            DisplayMaterialUsingAttach(modifynum, totalprice)
        End If
        drL.Close()
        connL.Close()
    End Sub

    Sub DisplayMaterialUsingAttach(modifynum As Long, totalprice As Double)
        Dim connL1 As New SqlConnection
        Dim drL1 As SqlDataReader
        Dim i As Integer
        i = 4
        SqlCmd = "Select T0.itemcode,T0.itemname,T0.quantity,T0.price,T0.method,T0.comment,T0.num " &
        "FROM [dbo].[@XSMLS] T0 " &
        "WHERE T0.head=0 And T0.[docentry] =" & docnum
        drL1 = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL1)
        If (drL1.HasRows) Then
            AddT.Enabled = False
            drL1.Read()
            ContentT.Rows(i).Cells(0).Text = i - 3
            ContentT.Rows(i).Cells(1).Text = drL1(0)
            ContentT.Rows(i).Cells(2).Text = drL1(1)
            ContentT.Rows(i).Cells(3).Text = drL1(2)
            ContentT.Rows(i).Cells(4).Text = Format(drL1(3), "###,###.##")
            ContentT.Rows(i).Cells(5).Text = Format(drL1(3) * drL1(2), "###,###.##") '總價
            ContentT.Rows(i).Cells(12).Text = drL1(4) '報廢原因
            ContentT.Rows(i).Cells(13).Text = drL1(5) '備註
            If (docstatus = "B" Or docstatus = "E" Or docstatus = "D" Or docstatus = "R" Or docstatus = "A") Then
                If (CType(ContentT.FindControl("chkdel_" & i & "_9"), CheckBox).Checked) Then
                    CType(ContentT.FindControl("hypermodify_" & i & "_8"), HyperLink).NavigateUrl = "cLsignoff.aspx?smid=sg&smode=2&act=material_del&status=" & docstatus &
                "&docnum=" & docnum & "&formstatusindex=" & formstatusindex &
                "&formtypeindex=" & formtypeindex & "&sfid=" & sfid & "&subject=" & TxtSubject.Text & "&num=" & drL1(6)
                    CType(ContentT.FindControl("hypermodify_" & i & "_8"), HyperLink).Text = "刪除"
                    CType(ContentT.FindControl("hypermodify_" & i & "_8"), HyperLink).Enabled = True
                Else
                    If (drL1(6) <> modifynum) Then
                        CType(ContentT.FindControl("hypermodify_" & i & "_8"), HyperLink).NavigateUrl = "cLsignoff.aspx?smid=sg&smode=2&act=material_modify&status=" & docstatus &
                    "&docnum=" & docnum & "&formstatusindex=" & formstatusindex &
                    "&formtypeindex=" & formtypeindex & "&sfid=" & sfid & "&subject=" & TxtSubject.Text & "&num=" & drL1(6)
                        CType(ContentT.FindControl("hypermodify_" & i & "_8"), HyperLink).Text = "修改"
                        CType(ContentT.FindControl("hypermodify_" & i & "_8"), HyperLink).Enabled = True
                    Else
                        CType(ContentT.FindControl("hypermodify_" & i & "_8"), HyperLink).Text = "修改中.."
                        CType(ContentT.FindControl("hypermodify_" & i & "_8"), HyperLink).Enabled = False
                        AddT.Enabled = True
                        TxtItemcode.Enabled = False
                        TxtItemname.Enabled = False
                        TxtQty.Enabled = False
                        ChkUsingAttach.Enabled = False
                    End If
                    'CType(ContentT.FindControl("hypermodify_" & i & "_8"), HyperLink).Text = "NA"
                    'CType(ContentT.FindControl("hypermodify_" & i & "_8"), HyperLink).Enabled = False
                End If
            End If
            ContentT.Rows(i + 1).Cells(4).Text = "Total"
            ContentT.Rows(i + 1).Cells(5).Text = Format(totalprice, "###,###.##")
            TxtPrice.Text = Format(totalprice, "###,###.##")
        End If
        drL1.Close()
        connL1.Close()
    End Sub
    Sub ShowMaterialList101(maindoc As Long)
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim i As Integer
        i = 2 '料件在Table 之起始列
        SqlCmd = "Select T0.itemcode,T0.itemname,T0.quantity,T0.rtnqty,T0.method,T0.comment,T0.num " &
                "FROM [dbo].[@XSMLS] T0 " &
                "WHERE T0.head=0 And T0.[docentry] =" & maindoc & " ORDER BY T0.num"
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            Do While (drL.Read())
                ContentT.Rows(i).Cells(0).Text = i - 1 '項次
                ContentT.Rows(i).Cells(1).Text = drL(0) '料號
                ContentT.Rows(i).Cells(2).Text = drL(1) '說明
                ContentT.Rows(i).Cells(3).Text = drL(2) '離倉數量
                ContentT.Rows(i).Cells(4).Text = drL(3) '已還數量
                'ContentT.Rows(i).Cells(5).Text = 0 '此次還回數量
                'CType(ContentT.FindControl("txtreturn_" & i & "_5"), TextBox).Text =
                ContentT.Rows(i).Cells(6).Text = drL(4) '當初離倉原因
                ContentT.Rows(i).Cells(7).Text = drL(5) '備註
                i = i + 1 'qqqqq
            Loop
        End If
        drL.Close()
        connL.Close()
    End Sub
    Sub ShowMaterialList23_24(modifynum As Long)
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim i As Integer
        Dim descrip As String = ""
        SqlCmd = "Select descrip FROM [dbo].[@XSMLS] T0 WHERE T0.[docentry] =" & docnum & " and head=1"
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            drL.Read()
            descrip = drL(0)
        End If
        drL.Close()
        connL.Close()
        TxtReason.Text = descrip
        If (Not IsPostBack) Then
            If (act = "notsave" And (sfid = 23 Or sfid = 24)) Then
                If (TxtReason.Text.Length <= 60) Then
                    CommUtil.ShowMsg(Me, "已儲存,但還有欄位未填入或沒附檔(或料件)" &
                                     "另需填寫借料人/借料單位/借料用途")
                Else
                    CommUtil.ShowMsg(Me, "已儲存,但還有欄位未填入或沒附檔(或料件)")
                End If
            End If
        End If
        If ((System.Text.RegularExpressions.Regex.Matches(descrip, "\r\n").Count + 1) > 7) Then
            TxtReason.Rows = System.Text.RegularExpressions.Regex.Matches(descrip, "\r\n").Count + 1
        End If
        i = 4 '料件在Table 之起始列
        SqlCmd = "Select T0.itemcode,T0.itemname,T0.quantity,T0.price,T0.method,T0.comment,T0.num,T0.rtnqty " &
                "FROM [dbo].[@XSMLS] T0 " &
                "WHERE T0.head=0 And T0.[docentry] =" & docnum & " ORDER BY T0.num"
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            Do While (drL.Read())
                ContentT.Rows(i).Cells(0).Text = i - 3 '項次
                ContentT.Rows(i).Cells(1).Text = drL(0) '料號
                ContentT.Rows(i).Cells(2).Text = drL(1) '說明
                ContentT.Rows(i).Cells(3).Text = drL(2) '需求數量
                ContentT.Rows(i).Cells(4).Text = drL(7) '還回數量
                ContentT.Rows(i).Cells(5).Text = drL(4) '報廢原因
                ContentT.Rows(i).Cells(6).Text = drL(5) '備註
                If (docstatus = "B" Or docstatus = "E" Or docstatus = "D" Or docstatus = "R" Or docstatus = "A") Then
                    If (CType(ContentT.FindControl("chkdel_" & i & "_9"), CheckBox).Checked) Then
                        CType(ContentT.FindControl("hypermodify_" & i & "_8"), HyperLink).NavigateUrl = "cLsignoff.aspx?smid=sg&smode=2&act=material_del&status=" & docstatus &
                    "&docnum=" & docnum & "&formstatusindex=" & formstatusindex &
                    "&formtypeindex=" & formtypeindex & "&sfid=" & sfid & "&subject=" & TxtSubject.Text & "&num=" & drL(6)
                        CType(ContentT.FindControl("hypermodify_" & i & "_8"), HyperLink).Text = "刪除"
                    Else
                        If (drL(6) <> modifynum) Then
                            CType(ContentT.FindControl("hypermodify_" & i & "_8"), HyperLink).NavigateUrl = "cLsignoff.aspx?smid=sg&smode=2&act=material_modify&status=" & docstatus &
                        "&docnum=" & docnum & "&formstatusindex=" & formstatusindex &
                        "&formtypeindex=" & formtypeindex & "&sfid=" & sfid & "&subject=" & TxtSubject.Text & "&num=" & drL(6)
                            CType(ContentT.FindControl("hypermodify_" & i & "_8"), HyperLink).Text = "修改"
                            CType(ContentT.FindControl("hypermodify_" & i & "_8"), HyperLink).Enabled = True
                        Else
                            CType(ContentT.FindControl("hypermodify_" & i & "_8"), HyperLink).Text = "修改中.."
                            CType(ContentT.FindControl("hypermodify_" & i & "_8"), HyperLink).Enabled = False
                        End If
                    End If
                End If
                i = i + 1
            Loop
        End If
        drL.Close()
        connL.Close()
    End Sub

    Sub ShowXRSCT()
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim str() As String
        Dim str1() As String
        Dim BColor As Drawing.Color
        BColor = System.Drawing.Color.LightBlue
        SqlCmd = "Select T0.idname,convert(Char(12),T0.CDate,111) As CDate,IsNull(convert(Char(12),T0.albdate,111),'N') as albdate, " &
                     "T0.albhour,T0.albmin,IsNull(convert(char(12),T0.aledate,111),'N') as aledate, " &
                     "T0.alehour,T0.alemin,T0.createname,T0.rsreason,T0.id,T0.createid,T0.v5id,T0.id " &
                     "FROM [dbo].[@XRSCT] T0 WHERE T0.[docentry] =" & docnum
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            drL.Read()
            str = Split(drL(1), "/")
            ContentT.Rows(0).Cells(1).Text = "&nbsp&nbsp" & str(0) & " 年 " & str(1) & " 月 " & str(2) & " 日"
            If (docstatus = "A" Or docstatus = "E" Or docstatus = "D" Or docstatus = "B" Or docstatus = "R") Then
                If (Trim(drL(2)) <> "N") Then
                    str = Split(drL(2), "/")
                    'MsgBox(drL(2) & " year:" & str(0) & " month:" & str(1) & " day:" & str(2))
                    CType(ContentT.FindControl("ddlfromyear_2_1"), DropDownList).SelectedValue = str(0)
                    CType(ContentT.FindControl("ddlfrommonth_2_1"), DropDownList).SelectedValue = str(1)
                    CType(ContentT.FindControl("ddlfromday_2_1"), DropDownList).SelectedValue = Trim(str(2))
                Else
                    CType(ContentT.FindControl("ddlfromyear_2_1"), DropDownList).SelectedValue = Now.Year
                    If (CInt(Now.Month) >= 10) Then
                        CType(ContentT.FindControl("ddlfrommonth_2_1"), DropDownList).SelectedValue = Now.Month
                    Else
                        CType(ContentT.FindControl("ddlfrommonth_2_1"), DropDownList).SelectedValue = "0" & Now.Month
                    End If
                    CType(ContentT.FindControl("ddlfromday_2_1"), DropDownList).SelectedIndex = 0
                End If
                If (Trim(drL(5)) <> "N") Then
                    str = Split(drL(5), "/")
                    CType(ContentT.FindControl("ddltoyear_2_1"), DropDownList).SelectedValue = str(0)
                    CType(ContentT.FindControl("ddltomonth_2_1"), DropDownList).SelectedValue = str(1)
                    CType(ContentT.FindControl("ddltoday_2_1"), DropDownList).SelectedValue = Trim(str(2))
                Else
                    CType(ContentT.FindControl("ddltoyear_2_1"), DropDownList).SelectedValue = Now.Year
                    If (CInt(Now.Month) >= 10) Then
                        CType(ContentT.FindControl("ddltomonth_2_1"), DropDownList).SelectedValue = Now.Month
                    Else
                        CType(ContentT.FindControl("ddltomonth_2_1"), DropDownList).SelectedValue = "0" & Now.Month
                    End If
                End If
                CType(ContentT.FindControl("ddlfromhour_3_1"), DropDownList).SelectedValue = drL(3).ToString.PadLeft(2, "0")
                CType(ContentT.FindControl("ddlfrommin_3_1"), DropDownList).SelectedValue = drL(4).ToString.PadLeft(2, "0")
                CType(ContentT.FindControl("ddltohour_3_1"), DropDownList).SelectedValue = drL(6).ToString.PadLeft(2, "0")
                CType(ContentT.FindControl("ddltomin_3_1"), DropDownList).SelectedValue = drL(7).ToString.PadLeft(2, "0")

                CType(ContentT.FindControl("txtrsreason_4_1"), TextBox).Text = drL(9)

                CType(ContentT.FindControl("ddlrsuser_6_0"), DropDownList).SelectedValue = drL(13) & " " & drL(0)
                'ContentT.Rows(6).Cells(0).Text = drL(13)
                CType(ContentT.FindControl("txtv5usercode_6_1"), TextBox).Text = drL(12)
            Else
                str = Split(drL(2), "/")
                str1 = Split(drL(5), "/")
                ContentT.Rows(1).Cells(1).Text = str(0) & "&nbsp&nbsp年&nbsp&nbsp" & str(1) & "&nbsp&nbsp月&nbsp&nbsp" & str(2) &
                                                "&nbsp&nbsp日&nbsp&nbsp&nbsp&nbsp至&nbsp&nbsp&nbsp&nbsp" & str1(0) & "&nbsp&nbsp年&nbsp&nbsp" &
                                                str1(1) & "&nbsp&nbsp月&nbsp&nbsp" & str1(2) & "&nbsp&nbsp日&nbsp&nbsp&nbsp&nbsp止"
                ContentT.Rows(2).Cells(1).Text = "&nbsp&nbsp起始&nbsp&nbsp" & drL(3).ToString.PadLeft(2, "0") & "&nbsp&nbsp時&nbsp&nbsp" & drL(4).ToString.PadLeft(2, "0") &
                                                "&nbsp&nbsp分&nbsp&nbsp&nbsp&nbsp至&nbsp&nbsp&nbsp&nbsp" & drL(6).ToString.PadLeft(2, "0") & "&nbsp&nbsp時&nbsp&nbsp" &
                                                drL(7).ToString.PadLeft(2, "0") & "&nbsp&nbsp分&nbsp&nbsp&nbsp&nbsp止&nbsp&nbsp&nbsp&nbsp"
                ContentT.Rows(3).Cells(1).Text = drL(9)
                ContentT.Rows(5).Cells(0).Text = drL(0)
                ContentT.Rows(5).Cells(1).Text = drL(12)
            End If
        End If
        drL.Close()
        connL.Close()
    End Sub
    Sub ShowXCMRT()
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        SqlCmd = "Select T0.reportdate,T0.machinetype,T0.cusname,cusfactoryOrmo,model,machineserialOrwo,installdateOrshipdate, " &
                     "problemtype,typedescrip,verandspec,faeperson, " &
                     "problemdescrip,processdescrip,verifydescrip,problemnote,firstinstallOrnoassign,inwarranty,qcperson " &
                     "FROM [dbo].[@XCMRT] T0 WHERE T0.[docentry] =" & docnum
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)

        If (drL.HasRows) Then
            drL.Read()
            CType(ContentT.FindControl("txt_reportdate"), TextBox).Text = drL(0)
            CType(ContentT.FindControl("txt_machinetype"), TextBox).Text = drL(1)
            CType(ContentT.FindControl("txt_cusname"), TextBox).Text = drL(2)
            CType(ContentT.FindControl("txt_cusfactoryOrmo"), TextBox).Text = drL(3)
            CType(ContentT.FindControl("txt_model"), TextBox).Text = drL(4)
            CType(ContentT.FindControl("txt_machineserialOrwo"), TextBox).Text = drL(5)
            CType(ContentT.FindControl("txt_installdateOrshipdate"), TextBox).Text = drL(6)
            CType(ContentT.FindControl("txt_problemtype"), TextBox).Text = drL(7)
            CType(ContentT.FindControl("txt_typedescrip"), TextBox).Text = drL(8)
            CType(ContentT.FindControl("txt_verandspec"), TextBox).Text = drL(9)
            If (sfid = 22) Then
                CType(ContentT.FindControl("txt_faeperson"), TextBox).Text = drL(10)
                If (drL(15) = 1) Then
                    CType(ContentT.FindControl("chk_firstinstallOrnoassign"), CheckBox).Checked = True
                End If
                If (drL(16) = 1) Then
                    CType(ContentT.FindControl("chk_inwarranty"), CheckBox).Checked = True
                End If
            ElseIf (sfid = 3) Then
                CType(ContentT.FindControl("txt_qcperson"), TextBox).Text = drL(17)
                If (drL(15) = 1) Then
                    CType(ContentT.FindControl("chk_firstinstallOrnoassign"), CheckBox).Checked = True
                End If
            End If
            CType(ContentT.FindControl("txt_problemdescrip"), TextBox).Text = drL(11)
            CType(ContentT.FindControl("txt_processdescrip"), TextBox).Text = drL(12)
            CType(ContentT.FindControl("txt_verifydescrip"), TextBox).Text = drL(13)
            CType(ContentT.FindControl("txt_problemnote"), TextBox).Text = drL(14)

            If ((System.Text.RegularExpressions.Regex.Matches(drL(11), "\r\n").Count + 1) <= 5) Then
                CType(ContentT.FindControl("txt_problemdescrip"), TextBox).Rows = 5
                'MsgBox(CType(ContentT.FindControl("txt_problemdescrip"), TextBox).Text.Length)
            Else
                CType(ContentT.FindControl("txt_problemdescrip"), TextBox).Rows = System.Text.RegularExpressions.Regex.Matches(drL(11), "\r\n").Count + 1
            End If

            If ((System.Text.RegularExpressions.Regex.Matches(drL(12), "\r\n").Count + 1) <= 5) Then
                CType(ContentT.FindControl("txt_processdescrip"), TextBox).Rows = 5
            Else
                CType(ContentT.FindControl("txt_processdescrip"), TextBox).Rows = System.Text.RegularExpressions.Regex.Matches(drL(12), "\r\n").Count + 1
            End If

            If ((System.Text.RegularExpressions.Regex.Matches(drL(13), "\r\n").Count + 1) <= 5) Then
                CType(ContentT.FindControl("txt_verifydescrip"), TextBox).Rows = 5
            Else
                CType(ContentT.FindControl("txt_verifydescrip"), TextBox).Rows = System.Text.RegularExpressions.Regex.Matches(drL(13), "\r\n").Count + 1
            End If

            If ((System.Text.RegularExpressions.Regex.Matches(drL(14), "\r\n").Count + 1) <= 5) Then
                CType(ContentT.FindControl("txt_problemnote"), TextBox).Rows = 5
            Else
                CType(ContentT.FindControl("txt_problemnote"), TextBox).Rows = System.Text.RegularExpressions.Regex.Matches(drL(14), "\r\n").Count + 1
            End If
        End If
        drL.Close()
        connL.Close()
    End Sub
    'Sub ShowXFMRT()
    '    Dim connL As New SqlConnection
    '    Dim drL As SqlDataReader
    '    SqlCmd = "Select T0.reportdate,T0.machinetype,T0.cusname,mo,model,wo,shipdate, " &
    '                 "problemtype,typedescrip,verandspec,qcperson, " &
    '                 "problemdescrip,processdescrip,verifydescrip,problemnote,noassign " &
    '                 "FROM [dbo].[@XFMRT] T0 WHERE T0.[docentry] =" & docnum
    '    drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)

    '    If (drL.HasRows) Then
    '        drL.Read()
    '        CType(ContentT.FindControl("txt_reportdate"), TextBox).Text = drL(0)
    '        CType(ContentT.FindControl("txt_machinetype"), TextBox).Text = drL(1)
    '        CType(ContentT.FindControl("txt_cusname"), TextBox).Text = drL(2)
    '        CType(ContentT.FindControl("txt_mo"), TextBox).Text = drL(3)
    '        CType(ContentT.FindControl("txt_model"), TextBox).Text = drL(4)
    '        CType(ContentT.FindControl("txt_wo"), TextBox).Text = drL(5)
    '        CType(ContentT.FindControl("txt_shipdate"), TextBox).Text = drL(6)
    '        CType(ContentT.FindControl("txt_problemtype"), TextBox).Text = drL(7)
    '        CType(ContentT.FindControl("txt_typedescrip"), TextBox).Text = drL(8)
    '        CType(ContentT.FindControl("txt_verandspec"), TextBox).Text = drL(9)
    '        CType(ContentT.FindControl("txt_qcperson"), TextBox).Text = drL(10)
    '        CType(ContentT.FindControl("txt_problemdescrip"), TextBox).Text = drL(11)
    '        CType(ContentT.FindControl("txt_processdescrip"), TextBox).Text = drL(12)
    '        CType(ContentT.FindControl("txt_verifydescrip"), TextBox).Text = drL(13)
    '        CType(ContentT.FindControl("txt_problemnote"), TextBox).Text = drL(14)
    '        If (drL(15) = 1) Then
    '            CType(ContentT.FindControl("chk_noassign"), CheckBox).Checked = True
    '        End If
    '        If ((System.Text.RegularExpressions.Regex.Matches(drL(11), "\r\n").Count + 1) <= 5) Then
    '            CType(ContentT.FindControl("txt_problemdescrip"), TextBox).Rows = 5
    '            'MsgBox(CType(ContentT.FindControl("txt_problemdescrip"), TextBox).Text.Length)
    '        Else
    '            CType(ContentT.FindControl("txt_problemdescrip"), TextBox).Rows = System.Text.RegularExpressions.Regex.Matches(drL(11), "\r\n").Count + 1
    '        End If

    '        If ((System.Text.RegularExpressions.Regex.Matches(drL(12), "\r\n").Count + 1) <= 5) Then
    '            CType(ContentT.FindControl("txt_processdescrip"), TextBox).Rows = 5
    '        Else
    '            CType(ContentT.FindControl("txt_processdescrip"), TextBox).Rows = System.Text.RegularExpressions.Regex.Matches(drL(12), "\r\n").Count + 1
    '        End If

    '        If ((System.Text.RegularExpressions.Regex.Matches(drL(13), "\r\n").Count + 1) <= 5) Then
    '            CType(ContentT.FindControl("txt_verifydescrip"), TextBox).Rows = 5
    '        Else
    '            CType(ContentT.FindControl("txt_verifydescrip"), TextBox).Rows = System.Text.RegularExpressions.Regex.Matches(drL(13), "\r\n").Count + 1
    '        End If

    '        If ((System.Text.RegularExpressions.Regex.Matches(drL(14), "\r\n").Count + 1) <= 5) Then
    '            CType(ContentT.FindControl("txt_problemnote"), TextBox).Rows = 5
    '        Else
    '            CType(ContentT.FindControl("txt_problemnote"), TextBox).Rows = System.Text.RegularExpressions.Regex.Matches(drL(14), "\r\n").Count + 1
    '        End If
    '    End If
    '    drL.Close()
    '    connL.Close()
    'End Sub
    Sub ShowXGCT()
        Dim connL, connL2 As New SqlConnection
        Dim drL, drL2 As SqlDataReader
        Dim BColor As Drawing.Color
        Dim senddate, sendperson, subject As String
        Dim nowdate As String
        BColor = System.Drawing.Color.LightBlue
        senddate = ""
        sendperson = ""
        subject = ""
        SqlCmd = "Select convert(char(12),T0.docdate,111) ,sname,subject from [dbo].[@XASCH] T0 " &
        "where docnum=" & docnum
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            drL.Read()
            senddate = drL(0)
            sendperson = drL(1)
            subject = drL(2)
        End If
        drL.Close()
        connL.Close()

        SqlCmd = "Select T0.ctdept,T0.ctperson,T0.ctdescrip " &
                     "FROM [dbo].[@XGCT] T0 WHERE T0.[docentry] =" & docnum
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)

        If (drL.HasRows) Then
            drL.Read()
            TxtDept.Text = drL(0)
            TxtPerson.Text = drL(1)
            TxtReason.Text = drL(2)
            If (docstatus = "A" Or docstatus = "E" Or docstatus = "D") Then

            Else
                nowdate = Format(Now(), "yyyy/MM/dd")
                SqlCmd = "Select IsNull(convert(char(12),T0.signdate,111),'M'),T0.uname " &
                        "FROM [dbo].[@XSPWT] T0 WHERE T0.[docentry] =" & docnum & "and seq=1"
                drL2 = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL2)
                If (drL2.HasRows) Then
                    drL2.Read()
                    If (Trim(drL2(0)) <> "M") Then '要加Trim才行
                        senddate = drL2(0)
                    Else
                        senddate = nowdate
                    End If
                    sendperson = drL2(1)
                End If
                drL2.Close()
                connL2.Close()
            End If
            If ((System.Text.RegularExpressions.Regex.Matches(drL(2), "\r\n").Count + 1) <= 6) Then
                TxtReason.Rows = 6
            Else
                TxtReason.Rows = System.Text.RegularExpressions.Regex.Matches(drL(2), "\r\n").Count + 1
            End If
        End If
        ContentT.Rows(1).Cells(1).Text = senddate
        ContentT.Rows(1).Cells(3).Text = sendperson
        ContentT.Rows(3).Cells(1).Text = subject
        drL.Close()
        connL.Close()
    End Sub
    Sub ShowXMSCT()
        Dim connL, connL2 As New SqlConnection
        Dim drL, drL2 As SqlDataReader
        Dim ddl_model As String
        Dim BColor As Drawing.Color
        BColor = System.Drawing.Color.LightBlue
        SqlCmd = "Select rbl_area, txt_amount, rbl_plheight, rbl_withfixture, rbl_pcbdir, rbl_oslang, rbl_camerapixel, rbl_rgb, " &
                         "rbl_resolution, rbl_rbgcontrol, rbl_coaxialinstall, rbl_coaxialcolor, rbl_belttype, chk_flux, chk_anti, " &
                         "txt_sales, ddl_model, txt_customer, txt_shipmodel, txt_shipdate, txt_uutlength, txt_uutwidth, txt_uutweight, " &
                         "txt_uutthick, txt_plotherheight, txt_plotherheighttol, txt_fixturesize, txt_pcbsizeX, txt_pcbsizeY, txt_cycletime, " &
                         "txt_zmm, txt_topspace, txt_botspace, txt_otherresolution, txt_memo,txt_dzmm,rbl_tblens,chk_sidecamera,chk_upz,chk_downz " &
                     "FROM [dbo].[@XMSCT] T0 WHERE T0.[docentry] =" & docnum
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            drL.Read()
            If (drL(16) <> "請選擇") Then
                SqlCmd = "SELECT T0.u_mdesc " &
                     "FROM dbo.[@UMMD] T0 where T0.u_model='" & drL(16) & "'"
                drL2 = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL2)
                If (drL2.HasRows) Then
                    drL2.Read()
                    ddl_model = drL(16) & "-" & drL2(0)
                Else
                    ddl_model = drL(16)
                End If

                drL2.Close()
                connL2.Close()
            Else
                ddl_model = drL(16)
            End If
            CType(ContentT.FindControl("rbl_area"), RadioButtonList).SelectedIndex = drL(0)
            If (drL(1) <> 0) Then
                CType(ContentT.FindControl("txt_amount"), TextBox).Text = drL(1)
            Else
                CType(ContentT.FindControl("txt_amount"), TextBox).Text = ""
            End If
            CType(ContentT.FindControl("rbl_plheight"), RadioButtonList).SelectedIndex = drL(2)
            CType(ContentT.FindControl("rbl_withfixture"), RadioButtonList).SelectedIndex = drL(3)
            CType(ContentT.FindControl("rbl_pcbdir"), RadioButtonList).SelectedIndex = drL(4)
            CType(ContentT.FindControl("rbl_oslang"), RadioButtonList).SelectedIndex = drL(5)
            CType(ContentT.FindControl("rbl_camerapixel"), RadioButtonList).SelectedIndex = drL(6)
            CType(ContentT.FindControl("rbl_rgb"), RadioButtonList).SelectedIndex = drL(7)
            CType(ContentT.FindControl("rbl_resolution"), RadioButtonList).SelectedIndex = drL(8)
            CType(ContentT.FindControl("rbl_rbgcontrol"), RadioButtonList).SelectedIndex = drL(9)
            CType(ContentT.FindControl("rbl_coaxialinstall"), RadioButtonList).SelectedIndex = drL(10)
            CType(ContentT.FindControl("rbl_coaxialcolor"), RadioButtonList).SelectedIndex = drL(11)
            CType(ContentT.FindControl("rbl_belttype"), RadioButtonList).SelectedIndex = drL(12)
            CType(ContentT.FindControl("chk_flux"), CheckBox).Checked = drL(13)
            CType(ContentT.FindControl("chk_anti"), CheckBox).Checked = drL(14)

            CType(ContentT.FindControl("txt_sales"), TextBox).Text = drL(15)
            CType(ContentT.FindControl("ddl_model"), DropDownList).SelectedValue = ddl_model
            CType(ContentT.FindControl("txt_customer"), TextBox).Text = drL(17)
            CType(ContentT.FindControl("txt_shipmodel"), TextBox).Text = drL(18)
            CType(ContentT.FindControl("txt_shipdate"), TextBox).Text = drL(19)
            CType(ContentT.FindControl("txt_uutlength"), TextBox).Text = drL(20)
            CType(ContentT.FindControl("txt_uutwidth"), TextBox).Text = drL(21)
            CType(ContentT.FindControl("txt_uutweight"), TextBox).Text = drL(22)
            CType(ContentT.FindControl("txt_uutthick"), TextBox).Text = drL(23)
            CType(ContentT.FindControl("txt_plotherheight"), TextBox).Text = drL(24)
            CType(ContentT.FindControl("txt_plotherheighttol"), TextBox).Text = drL(25)
            CType(ContentT.FindControl("txt_fixturesize"), TextBox).Text = drL(26)
            CType(ContentT.FindControl("txt_pcbsizeX"), TextBox).Text = drL(27)
            CType(ContentT.FindControl("txt_pcbsizeY"), TextBox).Text = drL(28)
            CType(ContentT.FindControl("txt_cycletime"), TextBox).Text = drL(29)
            CType(ContentT.FindControl("txt_zmm"), TextBox).Text = drL(30)
            CType(ContentT.FindControl("txt_topspace"), TextBox).Text = drL(31)
            CType(ContentT.FindControl("txt_botspace"), TextBox).Text = drL(32)
            CType(ContentT.FindControl("txt_otherresolution"), TextBox).Text = drL(33)
            CType(ContentT.FindControl("txt_memo"), TextBox).Text = drL(34)
            If ((System.Text.RegularExpressions.Regex.Matches(drL(34), "\r\n").Count + 1) <= 6) Then
                CType(ContentT.FindControl("txt_memo"), TextBox).Rows = 6
            Else
                CType(ContentT.FindControl("txt_memo"), TextBox).Rows = System.Text.RegularExpressions.Regex.Matches(drL(34), "\r\n").Count + 1
            End If
            If (CType(ContentT.FindControl("rbl_plheight"), RadioButtonList).SelectedIndex = 1) Then
                CType(ContentT.FindControl("txt_plotherheight"), TextBox).BackColor = BColor 'System.Drawing.Color.LightGreen
                CType(ContentT.FindControl("txt_plotherheighttol"), TextBox).BackColor = BColor 'System.Drawing.Color.LightGreen
            End If
            If (CType(ContentT.FindControl("rbl_resolution"), RadioButtonList).SelectedIndex = 11) Then
                CType(ContentT.FindControl("txt_otherresolution"), TextBox).BackColor = BColor 'System.Drawing.Color.LightGreen
            End If
            CType(ContentT.FindControl("txt_dzmm"), TextBox).Text = drL(35)
            CType(ContentT.FindControl("rbl_tblens"), RadioButtonList).SelectedIndex = drL(36)
            CType(ContentT.FindControl("chk_sidecamera"), CheckBox).Checked = drL(37)
            CType(ContentT.FindControl("chk_upz"), CheckBox).Checked = drL(38)
            CType(ContentT.FindControl("chk_downz"), CheckBox).Checked = drL(39)
        End If
        drL.Close()
        connL.Close()
        'If (docstatus <> "A" And docstatus <> "E" And docstatus <> "D" And docstatus <> "R" And docstatus <> "B") Then
        'ContentT.Enabled = False
        'End If
    End Sub
    Protected Sub ChkMaterialDel_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)
        AddTFieldClear()
        If (sfid = 23 Or sfid = 24) Then
            ShowMaterialList23_24(0)
        Else
            ShowMaterialList(0)
        End If
    End Sub
    Protected Sub ChkUsingAttach_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)
        If (ChkUsingAttach.Checked) Then
            TxtItemcode.Text = "ALL"
            TxtQty.Text = 1
            TxtItemname.Text = "一批料件,詳附檔"
            TxtNote.Text = "料件太多,以附檔呈現"
        Else
            TxtItemcode.Text = ""
            TxtQty.Text = ""
            TxtItemname.Text = ""
            TxtNote.Text = ""
        End If
    End Sub
    Sub FT0Create()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim Hyper As HyperLink
        Dim Labelx As Label
        Dim sign_status, nowseq, innerloop As Integer
        Dim lastsignoff As Boolean
        lastsignoff = False
        sign_status = 0
        Dim signofftype, cc As Integer
        signofftype = 0
        SqlCmd = "Select count(*) from [dbo].[@XSPWT] where docentry=" & docnum & " And uid='" & Session("s_id") & "'"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        dr.Read()
        cc = dr(0)
        dr.Close()
        connsap.Close()
        SqlCmd = "Select status,signprop,seq,innerloop from [dbo].[@XSPWT] where docentry=" & docnum & " and uid='" & Session("s_id") & "'"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        dr.Read()
        If (cc <> 0) Then
            nowseq = dr(2)
            innerloop = dr(3)
        End If
        If (cc = 1) Then '所有在簽核表中沒有重覆id
            sign_status = dr(0)
            signofftype = dr(1)
        ElseIf (cc = 0) Then

        Else '審核者及歸檔者是同一個
            If (docstatus = "D" Or docstatus = "B" Or docstatus = "R") Then '是審核者
                If (dr(2) <> 1) Then '非審核者 , 故需再read得到審核者
                    dr.Read()
                End If
            ElseIf (docstatus = "O") Then
                'nothing
            ElseIf (docstatus = "F" Or docstatus = "T") Then '是歸檔者
                'If (dr(2) = 1) Then '非歸檔者 , 故需再read得到歸檔者
                dr.Read()
                'End If
            End If
            sign_status = dr(0)
            signofftype = dr(1)
        End If
        dr.Close()
        connsap.Close()
        If (docstatus = "D") Then
            sign_status = 1
        End If
        If (cc <> 0) Then
            SqlCmd = "Select max(seq) from [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            If (dr(0) = nowseq) Then
                lastsignoff = True
            End If
            dr.Close()
            connsap.Close()
        End If
        tRow = New TableRow()
        tCell = New TableCell()
        'tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Left
        Hyper = New HyperLink
        Hyper.ID = "hyper_back0"
        Hyper.Text = "回前頁"
        If (fromasp = "signofftodo") Then
            Hyper.NavigateUrl = "~/signoff/signofftodo.aspx?smid=sg&smode=6&inchargeindex=" & inchargeindex & "&traceindex=" & traceindex &
                            "&formtypeindex=" & formtypeindex & "&indexpage=" & indexpage & "&inchargeid=" & inchargeid
        Else
            If (act = "skip" Or act = "skipok" Or act = "frommanage") Then
                Hyper.NavigateUrl = "~/signoff/signoffvip.aspx?smid=sg&smode=5&signflowmode=" & signflowmode & "&formstatusindex=" & formstatusindex &
                                "&formtypeindex=" & formtypeindex & "&indexpage=" & indexpage
            Else
                Hyper.NavigateUrl = "~/signoff/signoff.aspx?smid=sg&smode=1&signflowmode=" & signflowmode & "&formstatusindex=" & formstatusindex &
                                "&formtypeindex=" & formtypeindex & "&indexpage=" & indexpage
            End If
        End If
        Hyper.Font.Underline = False
        tCell.Controls.Add(Hyper)
        If (actmode = "signoff" Or actmode = "recycle" Or actmode = "single_signoff") Then
            Hyper.Visible = False
        Else
            Hyper.Visible = True
        End If

        Labelx = New Label()
        Labelx.ID = "label_fileul0"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp選擇上傳之檔案"
        tCell.Controls.Add(Labelx)
        Dim FileUL As New FileUpload()
        FileUL.ID = "fileul_0"
        tCell.Controls.Add(FileUL)
        If (docstatus = "A" Or docstatus = "E" Or docstatus = "D" Or ((docstatus = "R" Or docstatus = "B") And formstatusindex = 0 And sign_status = 1)) Then
            Labelx.Visible = True
            FileUL.Visible = True
        Else
            Labelx.Visible = False
            FileUL.Visible = False
        End If
        FileUL.Visible = False 'Disable , 用FT_m 的

        Dim ChkDel As New CheckBox
        ChkDel.ID = "chk_del_0"
        ChkDel.Text = "刪除檔案"
        ChkDel.AutoPostBack = True
        AddHandler ChkDel.CheckedChanged, AddressOf ChkDel_CheckedChanged
        tCell.Controls.Add(ChkDel)
        ChkDel.Visible = False 'Disable , 用FT_m 的

        Labelx = New Label()
        Labelx.ID = "label_upfile0"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        Dim BtnFileAct As New Button
        BtnFileAct.ID = "btn_fileact_0"
        BtnFileAct.Text = "上傳"
        AddHandler BtnFileAct.Click, AddressOf BtnFileAct_Click
        tCell.Controls.Add(BtnFileAct)
        If (docstatus = "A" Or docstatus = "E" Or docstatus = "D" Or ((docstatus = "R" Or docstatus = "B") And formstatusindex = 0 And sign_status = 1)) Then
            Labelx.Visible = True
            ChkDel.Visible = True
            BtnFileAct.Visible = True
        Else
            Labelx.Visible = False
            ChkDel.Visible = False
            BtnFileAct.Visible = False
        End If
        BtnFileAct.Visible = False 'Disable , 用FT_m 的
        tRow.Cells.Add(tCell)

        tCell = New TableCell()
        'tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Right

        BtnSave = New Button
        BtnSave.ID = "btn_save0"
        BtnSave.Width = 60
        If (docstatus = "A") Then
            BtnSave.Text = "建單"
        Else
            BtnSave.Text = "儲存"
        End If
        AddHandler BtnSave.Click, AddressOf BtnSave_Click
        tCell.Controls.Add(BtnSave)
        Labelsend = New Label
        Labelsend.ID = "label_send0"
        Labelsend.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelsend)
        BtnSend = New Button
        BtnSend.ID = "btn_send0"
        BtnSend.Width = 60
        BtnSend.Text = "送審"
        'act = "edit"
        AddHandler BtnSend.Click, AddressOf BtnSend_Click
        tCell.Controls.Add(BtnSend)

        Labeldel = New Label
        Labeldel.ID = "label_del0"
        Labeldel.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labeldel)
        BtnDel = New Button
        BtnDel.ID = "btn_del0"
        BtnDel.Width = 60
        BtnDel.Text = "刪除"
        AddHandler BtnDel.Click, AddressOf BtnDel_Click
        tCell.Controls.Add(BtnDel)

        Labelapproval = New Label
        Labelapproval.ID = "label_approval0"
        Labelapproval.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelapproval)
        BtnApproval = New Button
        BtnApproval.ID = "btn_approval0"
        BtnApproval.Width = 60
        BtnApproval.Text = "核准"
        AddHandler BtnApproval.Click, AddressOf BtnApproval_Click
        tCell.Controls.Add(BtnApproval)

        Labelreject = New Label
        Labelreject.ID = "label_reject0"
        Labelreject.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelreject)
        BtnReject = New Button
        BtnReject.ID = "btn_reject0"
        BtnReject.Width = 60
        If (cc <> 0) Then
            If (innerloop = 1 Or lastsignoff = True) Then
                BtnReject.Text = "駁回"
            Else
                BtnReject.Text = "反對"
            End If
        Else
            BtnReject.Text = "駁回"
        End If
        AddHandler BtnReject.Click, AddressOf BtnReject_Click
        tCell.Controls.Add(BtnReject)

        Labelx = New Label
        Labelx.ID = "label_suspend0"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        BtnSuspend = New Button
        BtnSuspend.ID = "btn_suspend0"
        BtnSuspend.Width = 70
        BtnSuspend.Text = "暫不簽核"
        AddHandler BtnSuspend.Click, AddressOf BtnNext_Click
        tCell.Controls.Add(BtnSuspend)

        ChkReturn = New CheckBox
        ChkReturn.ID = "chk_return_0"
        ChkReturn.Text = "退回審核人"
        AddHandler ChkReturn.CheckedChanged, AddressOf ChkReturn_CheckedChanged
        tCell.Controls.Add(ChkReturn)
        If (docstatus = "F" Or docstatus = "T") Then
            ChkReturn.Visible = False
        Else
            If (cc <> 0) Then
                If (innerloop = 1 Or lastsignoff = True) Then
                    ChkReturn.Visible = False
                Else
                    ChkReturn.Visible = True
                End If
            Else
                ChkReturn.Visible = False
            End If
        End If
        Labelrecall = New Label
        Labelrecall.ID = "label_recall0"
        Labelrecall.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelrecall)
        BtnRecall = New Button
        BtnRecall.ID = "btn_recall0"
        BtnRecall.Width = 60
        BtnRecall.Text = "抽回"
        AddHandler BtnRecall.Click, AddressOf BtnRecall_Click
        tCell.Controls.Add(BtnRecall)

        Labelcancel = New Label
        Labelcancel.ID = "label_cancel0"
        Labelcancel.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelcancel)
        BtnCancel = New Button
        BtnCancel.ID = "btn_cancel0"
        BtnCancel.Width = 60
        BtnCancel.Text = "作廢"
        AddHandler BtnCancel.Click, AddressOf BtnCancel_Click
        tCell.Controls.Add(BtnCancel)

        Labelarchieve = New Label
        Labelarchieve.ID = "label_archieve0"
        Labelarchieve.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelarchieve)
        BtnArchieve = New Button
        BtnArchieve.ID = "btn_archieve0"
        BtnArchieve.Width = 60
        BtnArchieve.Text = "歸檔"
        AddHandler BtnArchieve.Click, AddressOf BtnArchieve_Click
        tCell.Controls.Add(BtnArchieve)

        LabelBeInformed = New Label
        LabelBeInformed.ID = "label_BeInformed0"
        LabelBeInformed.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(LabelBeInformed)
        BtnBeInformed = New Button
        BtnBeInformed.ID = "btn_BeInformed0"
        BtnBeInformed.Width = 60
        BtnBeInformed.Text = "已知悉"
        AddHandler BtnBeInformed.Click, AddressOf BtnBeInformed_Click
        tCell.Controls.Add(BtnBeInformed)

        LabelSkip = New Label
        LabelSkip.ID = "label_skip0"
        LabelSkip.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(LabelSkip)
        BtnSkip = New Button
        BtnSkip.ID = "btn_skip0"
        BtnSkip.Width = 70
        BtnSkip.Text = "跳過簽核"
        AddHandler BtnSkip.Click, AddressOf BtnSkip_Click
        tCell.Controls.Add(BtnSkip)
        If (act = "skip") Then
            BtnSkip.Visible = True
        Else
            BtnSkip.Visible = False
        End If

        LabelLast = New Label
        LabelLast.ID = "label_last0"
        LabelLast.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(LabelLast)
        BtnLast = New Button
        BtnLast.ID = "btn_last0"
        BtnLast.Width = 60
        BtnLast.Text = "上一筆"
        AddHandler BtnLast.Click, AddressOf BtnLast_Click
        tCell.Controls.Add(BtnLast)

        LabelCount = New Label
        LabelCount.ID = "label_count0"
        LabelCount.Font.Bold = True
        If (signoffalready = False) Then '如果點選之簽核通知單為已覆核過, 則不顯示筆數 
            If (actmode = "") Then
                LabelCount.Text = "&nbsp&nbsp第&nbsp&nbsp 1/1 &nbsp&nbsp筆"
            Else
                If (Session("sgcount") <> 0) Then
                    LabelCount.Text = "&nbsp&nbsp第&nbsp&nbsp" & ((Session("startindex") + 1) & "/" & Session("sgcount")) & "&nbsp&nbsp筆"
                Else
                    LabelCount.Text = "&nbsp&nbsp第&nbsp&nbsp" & (Session("startindex") + 1) & "/1&nbsp&nbsp筆"
                End If
            End If
        End If
        tCell.Controls.Add(LabelCount)

        LabelNext = New Label
        LabelNext.ID = "label_next0"
        LabelNext.Text = "&nbsp&nbsp"
        tCell.Controls.Add(LabelNext)
        BtnNext = New Button
        BtnNext.ID = "btn_next0"
        BtnNext.Width = 60
        BtnNext.Text = "下一筆"
        AddHandler BtnNext.Click, AddressOf BtnNext_Click
        tCell.Controls.Add(BtnNext)

        SignOffStatusLabel = New Label
        SignOffStatusLabel.ID = "label_signoffstatus0"
        If (signofftype <> 2) Then '若為未送審,則用XSPWT 一定是無法找到(要用XASCH),所以signofftype會為0(那剛好0是所要的,故就不再另外處理)
            SignOffStatusLabel.Text = info
        Else
            If (sign_status = 1) Then
                SignOffStatusLabel.Text = "因你被設定為此單知悉人,只能觀看"
            Else
                SignOffStatusLabel.Text = "你已知悉過此單"
            End If
        End If
        tCell.Controls.Add(SignOffStatusLabel)

        'LabelPdf = New Label
        'LabelPdf.ID = "label_pdf"
        'LabelPdf.Text = "&nbsp&nbsp"
        'tCell.Controls.Add(LabelPdf)
        'BtnPdf = New Button
        'BtnPdf.ID = "btn_pdf"
        'BtnPdf.Width = 60
        'BtnPdf.Text = "轉Pdf"
        'AddHandler BtnPdf.Click, AddressOf BtnPdf_Click
        'tCell.Controls.Add(BtnPdf)


        BtnSend.Visible = False
        BtnDel.Visible = False
        BtnSave.Visible = False
        BtnCancel.Visible = False
        BtnApproval.Visible = False
        BtnReject.Visible = False
        BtnRecall.Visible = False
        BtnArchieve.Visible = False
        Labelsend.Visible = False
        Labeldel.Visible = False
        Labelapproval.Visible = False
        Labelreject.Visible = False
        Labelrecall.Visible = False
        Labelcancel.Visible = False
        Labelarchieve.Visible = False
        SignOffStatusLabel.Visible = True
        BtnBeInformed.Visible = False
        LabelBeInformed.Visible = False
        If (docstatus = "D" Or docstatus = "A" Or docstatus = "E") Then
            If (docstatus = "D") Then
                BtnSend.Visible = True
                Labelsend.Visible = True
                ChkReturn.Visible = False
            End If
            If (docstatus <> "A") Then
                BtnDel.Visible = True
                Labeldel.Visible = True
            End If
            BtnSave.Visible = True
        End If
        If ((docstatus = "R" Or docstatus = "B") And formstatusindex = 0 And sign_status = 1) Then
            BtnSend.Visible = True
            Labelsend.Visible = True
            ChkReturn.Visible = False
            BtnSend.Text = "再送審"
            BtnCancel.Visible = True
            Labelcancel.Visible = True
            BtnSave.Visible = True
        End If
        If (docstatus = "O" And formstatusindex = 0 And sign_status = 1) Then
            BtnApproval.Visible = True
            Labelapproval.Visible = True
            BtnReject.Visible = True
            Labelreject.Visible = True
        End If
        If (docstatus = "F" And sign_status = 1 And signofftype = 1) Then
            BtnArchieve.Visible = True
            Labelarchieve.Visible = True
        End If
        If (signofftype = 2 And sign_status = 1) Then
            BtnBeInformed.Visible = True
            LabelBeInformed.Visible = True
        End If
        Dim nextseq, status As Integer
        If (docstatus <> "A" And docstatus <> "E" And docstatus <> "D" And docstatus <> "F") Then
            SqlCmd = "Select status,seq from [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum & " and uid='" & Session("s_id") & "'"
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                dr.Read()
                status = dr(0)
                nextseq = dr(1) + 1
            End If
            dr.Close()
            connsap.Close()
            If (status = 2 Or status = 100) Then
                SqlCmd = "Select status from [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum & " and seq=" & nextseq
                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                If (dr.HasRows) Then
                    dr.Read()
                    status = dr(0)
                    dr.Close()
                    connsap.Close()
                    If (status = 1) Then '我已核 , 但下一關還未核 , 故可抽回
                        BtnRecall.Visible = True
                        Labelrecall.Visible = True
                    End If
                End If
                dr.Close()
                connsap.Close()
            End If
        End If
        BtnNext.Visible = True
        BtnSuspend.Visible = True
        BtnLast.Visible = True
        LabelNext.Visible = True
        LabelLast.Visible = True
        LabelCount.Visible = True
        If (sign_status <> 1) Then
            SignOffStatusLabel.Visible = True
        End If
        'If ((actmode = "recycle" Or actmode = "signoff") And sign_status = 1) Then
        Dim mes(2) As String
        If ((actmode = "recycle" Or actmode = "signoff" Or actmode = "recycle_login" Or actmode = "signoff_login")) Then
            If ((Session("startindex") = (Session("sgcount") - 1)) Or Session("sgcount") = 0) Then
                BtnNext.Enabled = False
                BtnSuspend.Enabled = False
                'BtnLast.Enabled = True
            End If
            If (Session("startindex") = 0) Then
                'BtnNext.Enabled = True
                BtnLast.Enabled = False
            End If
            If (Session("sgcount") = 0) Then '已簽核過
                mes = CommSignOff.FormStatusMes(docnum, docstatus)
                SignOffStatusLabel.Text = mes(1)
                docstatus = mes(0)
                signoffalready = True
            End If
        Else
            BtnNext.Enabled = False
            BtnSuspend.Enabled = False
            BtnLast.Enabled = False
            'SignOffStatusLabel.Visible = True
        End If
        If (signoffalready = True) Then '如果點選之簽核通知單為已覆核過, 則不顯示
            BtnNext.Visible = False
            BtnSuspend.Visible = False
            BtnLast.Visible = False
            LabelNext.Visible = False
            LabelLast.Visible = False
            LabelCount.Visible = False
        End If
        tRow.Cells.Add(tCell)
        FT_0.Rows.Add(tRow)
        ViewState("docstatus") = docstatus
    End Sub
    Sub FT1Create()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim Hyper As HyperLink
        Dim Labelx As Label
        Dim sign_status, nowseq, innerloop As Integer
        Dim lastsignoff As Boolean
        lastsignoff = False
        sign_status = 0
        Dim signofftype, cc As Integer
        signofftype = 0
        SqlCmd = "Select count(*) from [dbo].[@XSPWT] where docentry=" & docnum & " and uid='" & Session("s_id") & "'"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        dr.Read()
        cc = dr(0)
        dr.Close()
        connsap.Close()
        SqlCmd = "Select status,signprop,seq,innerloop from [dbo].[@XSPWT] where docentry=" & docnum & " and uid='" & Session("s_id") & "'"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        dr.Read()
        If (cc <> 0) Then
            nowseq = dr(2)
            innerloop = dr(3)
        End If
        If (cc = 1) Then '所有在簽核表中沒有重覆id
            sign_status = dr(0)
            signofftype = dr(1)
        ElseIf (cc = 0) Then

        Else '審核者及歸檔者是同一個
            If (docstatus = "D" Or docstatus = "B" Or docstatus = "R") Then '是審核者
                If (dr(2) <> 1) Then '非審核者 , 故需再read得到審核者
                    dr.Read()
                End If
            ElseIf (docstatus = "O") Then
                'nothing
            ElseIf (docstatus = "F" Or docstatus = "T") Then '是歸檔者
                'If (dr(2) = 1) Then '非歸檔者 , 故需再read得到歸檔者
                dr.Read()
                'End If
            End If
            sign_status = dr(0)
            signofftype = dr(1)
        End If
        dr.Close()
        connsap.Close()
        If (docstatus = "D") Then
            sign_status = 1
        End If
        If (cc <> 0) Then
            SqlCmd = "Select max(seq) from [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            If (dr(0) = nowseq) Then
                lastsignoff = True
            End If
            dr.Close()
            connsap.Close()
        End If
        tRow = New TableRow()
        tCell = New TableCell()
        'tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Left
        Hyper = New HyperLink
        Hyper.ID = "hyper_back1"
        Hyper.Text = "回前頁"
        If (fromasp = "signofftodo") Then
            Hyper.NavigateUrl = "~/signoff/signofftodo.aspx?smid=sg&smode=6&inchargeindex=" & inchargeindex & "&traceindex=" & traceindex &
                            "&formtypeindex=" & formtypeindex & "&indexpage=" & indexpage & "&inchargeid=" & inchargeid
        Else
            If (act = "skip" Or act = "skipok" Or act = "frommanage") Then
                Hyper.NavigateUrl = "~/signoff/signoffvip.aspx?smid=sg&smode=5&signflowmode=" & signflowmode & "&formstatusindex=" & formstatusindex &
                                "&formtypeindex=" & formtypeindex & "&indexpage=" & indexpage
            Else
                Hyper.NavigateUrl = "~/signoff/signoff.aspx?smid=sg&smode=1&signflowmode=" & signflowmode & "&formstatusindex=" & formstatusindex &
                                "&formtypeindex=" & formtypeindex & "&indexpage=" & indexpage
            End If
        End If
        Hyper.Font.Underline = False
        tCell.Controls.Add(Hyper)
        If (actmode = "signoff" Or actmode = "recycle" Or actmode = "single_signoff") Then
            Hyper.Visible = False
        Else
            Hyper.Visible = True
        End If

        Labelx = New Label()
        Labelx.ID = "label_fileul1"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp選擇上傳之檔案"
        tCell.Controls.Add(Labelx)
        Dim FileUL As New FileUpload()
        FileUL.ID = "fileul_1"
        tCell.Controls.Add(FileUL)
        If (docstatus = "A" Or docstatus = "E" Or docstatus = "D" Or ((docstatus = "R" Or docstatus = "B") And formstatusindex = 0 And sign_status = 1)) Then
            Labelx.Visible = True
            FileUL.Visible = True
        Else
            Labelx.Visible = False
            FileUL.Visible = False
        End If
        Labelx.Visible = False
        FileUL.Visible = False '在此不用, 用 FT_m 的

        Dim ChkDel As New CheckBox
        ChkDel.ID = "chk_del_1"
        ChkDel.Text = "刪除檔案"
        ChkDel.AutoPostBack = True
        AddHandler ChkDel.CheckedChanged, AddressOf ChkDel_CheckedChanged
        tCell.Controls.Add(ChkDel)
        ChkDel.Visible = False '在此不用, 用 FT_m 的

        Labelx = New Label()
        Labelx.ID = "label_upfile1"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        Dim BtnFileAct As New Button
        BtnFileAct.ID = "btn_fileact_1"
        BtnFileAct.Text = "上傳"
        AddHandler BtnFileAct.Click, AddressOf BtnFileAct_Click
        tCell.Controls.Add(BtnFileAct)
        If (docstatus = "A" Or docstatus = "E" Or docstatus = "D" Or ((docstatus = "R" Or docstatus = "B") And formstatusindex = 0 And sign_status = 1)) Then
            Labelx.Visible = True
            ChkDel.Visible = True
            BtnFileAct.Visible = True
        Else
            Labelx.Visible = False
            ChkDel.Visible = False
            BtnFileAct.Visible = False
        End If
        ChkDel.Visible = False '在此不用, 用 FT_m 的
        BtnFileAct.Visible = False '在此不用, 用 FT_m 的
        tRow.Cells.Add(tCell)

        tCell = New TableCell()
        'tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Right

        BtnSave = New Button
        BtnSave.ID = "btn_save1"
        BtnSave.Width = 60
        If (docstatus = "A") Then
            BtnSave.Text = "建單"
        Else
            BtnSave.Text = "儲存"
        End If
        AddHandler BtnSave.Click, AddressOf BtnSave_Click
        tCell.Controls.Add(BtnSave)
        Labelsend = New Label
        Labelsend.ID = "label_send1"
        Labelsend.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelsend)
        BtnSend = New Button
        BtnSend.ID = "btn_send1"
        BtnSend.Width = 60
        BtnSend.Text = "送審"
        'act = "edit"
        AddHandler BtnSend.Click, AddressOf BtnSend_Click
        tCell.Controls.Add(BtnSend)

        Labeldel = New Label
        Labeldel.ID = "label_del1"
        Labeldel.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labeldel)
        BtnDel = New Button
        BtnDel.ID = "btn_del1"
        BtnDel.Width = 60
        BtnDel.Text = "刪除"
        AddHandler BtnDel.Click, AddressOf BtnDel_Click
        tCell.Controls.Add(BtnDel)

        Labelapproval = New Label
        Labelapproval.ID = "label_approval1"
        Labelapproval.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelapproval)
        BtnApproval = New Button
        BtnApproval.ID = "btn_approval1"
        BtnApproval.Width = 60
        BtnApproval.Text = "核准"
        AddHandler BtnApproval.Click, AddressOf BtnApproval_Click
        tCell.Controls.Add(BtnApproval)

        Labelreject = New Label
        Labelreject.ID = "label_reject1"
        Labelreject.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelreject)
        BtnReject = New Button
        BtnReject.ID = "btn_reject1"
        BtnReject.Width = 60
        If (cc <> 0) Then
            If (innerloop = 1 Or lastsignoff = True) Then
                BtnReject.Text = "駁回"
            Else
                BtnReject.Text = "反對"
            End If
        Else
            BtnReject.Text = "駁回"
        End If
        AddHandler BtnReject.Click, AddressOf BtnReject_Click
        tCell.Controls.Add(BtnReject)

        Labelx = New Label
        Labelx.ID = "label_suspend1"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        BtnSuspend = New Button
        BtnSuspend.ID = "btn_suspend1"
        BtnSuspend.Width = 70
        BtnSuspend.Text = "暫不簽核"
        AddHandler BtnSuspend.Click, AddressOf BtnNext_Click
        tCell.Controls.Add(BtnSuspend)

        ChkReturn = New CheckBox
        ChkReturn.ID = "chk_return_1"
        ChkReturn.Text = "退回審核人"
        ChkReturn.AutoPostBack = True
        AddHandler ChkReturn.CheckedChanged, AddressOf ChkReturn_CheckedChanged
        tCell.Controls.Add(ChkReturn)
        If (docstatus = "F" Or docstatus = "T") Then
            ChkReturn.Visible = False
        Else
            If (cc <> 0) Then
            If (innerloop = 1 Or lastsignoff = True) Then
                ChkReturn.Visible = False
            Else
                ChkReturn.Visible = True
            End If
        Else
            ChkReturn.Visible = False
        End If
        End If
        Labelrecall = New Label
        Labelrecall.ID = "label_recall1"
        Labelrecall.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelrecall)
        BtnRecall = New Button
        BtnRecall.ID = "btn_recall1"
        BtnRecall.Width = 60
        BtnRecall.Text = "抽回"
        AddHandler BtnRecall.Click, AddressOf BtnRecall_Click
        tCell.Controls.Add(BtnRecall)

        Labelcancel = New Label
        Labelcancel.ID = "label_cancel1"
        Labelcancel.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelcancel)
        BtnCancel = New Button
        BtnCancel.ID = "btn_cancel1"
        BtnCancel.Width = 60
        BtnCancel.Text = "作廢"
        AddHandler BtnCancel.Click, AddressOf BtnCancel_Click
        tCell.Controls.Add(BtnCancel)

        Labelarchieve = New Label
        Labelarchieve.ID = "label_archieve1"
        Labelarchieve.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelarchieve)
        BtnArchieve = New Button
        BtnArchieve.ID = "btn_archieve1"
        BtnArchieve.Width = 60
        BtnArchieve.Text = "歸檔"
        AddHandler BtnArchieve.Click, AddressOf BtnArchieve_Click
        tCell.Controls.Add(BtnArchieve)

        LabelBeInformed = New Label
        LabelBeInformed.ID = "label_BeInformed1"
        LabelBeInformed.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(LabelBeInformed)
        BtnBeInformed = New Button
        BtnBeInformed.ID = "btn_BeInformed1"
        BtnBeInformed.Width = 60
        BtnBeInformed.Text = "已知悉"
        AddHandler BtnBeInformed.Click, AddressOf BtnBeInformed_Click
        tCell.Controls.Add(BtnBeInformed)

        LabelSkip = New Label
        LabelSkip.ID = "label_skip1"
        LabelSkip.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(LabelSkip)
        BtnSkip = New Button
        BtnSkip.ID = "btn_skip1"
        BtnSkip.Width = 70
        BtnSkip.Text = "跳過簽核"
        AddHandler BtnSkip.Click, AddressOf BtnSkip_Click
        tCell.Controls.Add(BtnSkip)
        If (act = "skip") Then
            BtnSkip.Visible = True
        Else
            BtnSkip.Visible = False
        End If

        LabelLast = New Label
        LabelLast.ID = "label_last1"
        LabelLast.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(LabelLast)
        BtnLast = New Button
        BtnLast.ID = "btn_last1"
        BtnLast.Width = 60
        BtnLast.Text = "上一筆"
        AddHandler BtnLast.Click, AddressOf BtnLast_Click
        tCell.Controls.Add(BtnLast)

        LabelCount = New Label
        LabelCount.ID = "label_count1"
        LabelCount.Font.Bold = True
        If (signoffalready = False) Then '如果點選之簽核通知單為已覆核過, 則不顯示筆數 
            If (actmode = "") Then
                LabelCount.Text = "&nbsp&nbsp第&nbsp&nbsp 1/1 &nbsp&nbsp筆"
            Else
                If (Session("sgcount") <> 0) Then
                    LabelCount.Text = "&nbsp&nbsp第&nbsp&nbsp" & ((Session("startindex") + 1) & "/" & Session("sgcount")) & "&nbsp&nbsp筆"
                Else
                    LabelCount.Text = "&nbsp&nbsp第&nbsp&nbsp" & (Session("startindex") + 1) & "/1&nbsp&nbsp筆"
                End If
            End If
        End If
        tCell.Controls.Add(LabelCount)

        LabelNext = New Label
        LabelNext.ID = "label_next1"
        LabelNext.Text = "&nbsp&nbsp"
        tCell.Controls.Add(LabelNext)
        BtnNext = New Button
        BtnNext.ID = "btn_next1"
        BtnNext.Width = 60
        BtnNext.Text = "下一筆"
        AddHandler BtnNext.Click, AddressOf BtnNext_Click
        tCell.Controls.Add(BtnNext)

        SignOffStatusLabel = New Label
        SignOffStatusLabel.ID = "label_signoffstatus1"
        If (signofftype <> 2) Then '若為未送審,則用XSPWT 一定是無法找到(要用XASCH),所以signofftype會為0(那剛好0是所要的,故就不再另外處理)
            SignOffStatusLabel.Text = info
        Else
            If (sign_status = 1) Then
                SignOffStatusLabel.Text = "因你被設定為此單知悉人,只能觀看"
            Else
                SignOffStatusLabel.Text = "你已知悉過此單"
            End If
        End If
        tCell.Controls.Add(SignOffStatusLabel)

        'LabelPdf = New Label
        'LabelPdf.ID = "label_pdf"
        'LabelPdf.Text = "&nbsp&nbsp"
        'tCell.Controls.Add(LabelPdf)
        'BtnPdf = New Button
        'BtnPdf.ID = "btn_pdf"
        'BtnPdf.Width = 60
        'BtnPdf.Text = "轉Pdf"
        'AddHandler BtnPdf.Click, AddressOf BtnPdf_Click
        'tCell.Controls.Add(BtnPdf)


        BtnSend.Visible = False
        BtnDel.Visible = False
        BtnSave.Visible = False
        BtnCancel.Visible = False
        BtnApproval.Visible = False
        BtnReject.Visible = False
        BtnRecall.Visible = False
        BtnArchieve.Visible = False
        Labelsend.Visible = False
        Labeldel.Visible = False
        Labelapproval.Visible = False
        Labelreject.Visible = False
        Labelrecall.Visible = False
        Labelcancel.Visible = False
        Labelarchieve.Visible = False
        SignOffStatusLabel.Visible = True
        BtnBeInformed.Visible = False
        LabelBeInformed.Visible = False
        If (docstatus = "D" Or docstatus = "A" Or docstatus = "E") Then
            If (docstatus = "D") Then
                BtnSend.Visible = True
                Labelsend.Visible = True
                ChkReturn.Visible = False
            End If
            If (docstatus <> "A") Then
                BtnDel.Visible = True
                Labeldel.Visible = True
            End If
            BtnSave.Visible = True
        End If
        If ((docstatus = "R" Or docstatus = "B") And formstatusindex = 0 And sign_status = 1) Then
            BtnSend.Visible = True
            Labelsend.Visible = True
            ChkReturn.Visible = False
            BtnSend.Text = "再送審"
            BtnCancel.Visible = True
            Labelcancel.Visible = True
            BtnSave.Visible = True
        End If
        If (docstatus = "O" And formstatusindex = 0 And sign_status = 1) Then
            BtnApproval.Visible = True
            Labelapproval.Visible = True
            BtnReject.Visible = True
            Labelreject.Visible = True
        End If
        If (docstatus = "F" And sign_status = 1 And signofftype = 1) Then
            BtnArchieve.Visible = True
            Labelarchieve.Visible = True
        End If
        If (signofftype = 2 And sign_status = 1) Then
            BtnBeInformed.Visible = True
            LabelBeInformed.Visible = True
        End If
        Dim nextseq, status As Integer
        If (docstatus <> "A" And docstatus <> "E" And docstatus <> "D" And docstatus <> "F") Then
            SqlCmd = "Select status,seq from [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum & " and uid='" & Session("s_id") & "'"
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                dr.Read()
                status = dr(0)
                nextseq = dr(1) + 1
            End If
            dr.Close()
            connsap.Close()
            If (status = 2 Or status = 100) Then
                SqlCmd = "Select status from [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum & " and seq=" & nextseq
                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                If (dr.HasRows) Then
                    dr.Read()
                    status = dr(0)
                    dr.Close()
                    connsap.Close()
                    If (status = 1) Then '我已核 , 但下一關還未核 , 故可抽回
                        BtnRecall.Visible = True
                        Labelrecall.Visible = True
                    End If
                End If
                dr.Close()
                connsap.Close()
            End If
        End If
        BtnNext.Visible = True
        BtnSuspend.Visible = True
        BtnLast.Visible = True
        LabelNext.Visible = True
        LabelLast.Visible = True
        LabelCount.Visible = True
        If (sign_status <> 1) Then
            SignOffStatusLabel.Visible = True
        End If
        'If ((actmode = "recycle" Or actmode = "signoff") And sign_status = 1) Then
        Dim mes(2) As String
        If ((actmode = "recycle" Or actmode = "signoff" Or actmode = "recycle_login" Or actmode = "signoff_login")) Then
            If ((Session("startindex") = (Session("sgcount") - 1)) Or Session("sgcount") = 0) Then
                BtnNext.Enabled = False
                BtnSuspend.Enabled = False
                'BtnLast.Enabled = True
            End If
            If (Session("startindex") = 0) Then
                'BtnNext.Enabled = True
                BtnLast.Enabled = False
            End If
            If (Session("sgcount") = 0) Then '已簽核過
                mes = CommSignOff.FormStatusMes(docnum, docstatus)
                SignOffStatusLabel.Text = mes(1)
                docstatus = mes(0)
                signoffalready = True
            End If
        Else
            BtnNext.Enabled = False
            BtnSuspend.Enabled = False
            BtnLast.Enabled = False
            'SignOffStatusLabel.Visible = True
        End If
        If (signoffalready = True) Then '如果點選之簽核通知單為已覆核過, 則不顯示
            BtnNext.Visible = False
            BtnSuspend.Visible = False
            BtnLast.Visible = False
            LabelNext.Visible = False
            LabelLast.Visible = False
            LabelCount.Visible = False
        End If
        tRow.Cells.Add(tCell)
        FT_1.Rows.Add(tRow)
        ViewState("docstatus") = docstatus
    End Sub
    Function InitAddCommonHead(Now_time As String, sid As String, s_name As String, sfid As Integer)
        Dim num As Long
        Dim dept, area, deptcode, areacode, spos As String
        deptcode = ""
        areacode = ""
        area = ""
        dept = ""
        spos = ""
        SqlCmd = "select grp,branch,position from dbo.[user] where id='" & sid & "'"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        If (dr.HasRows) Then
            dr.Read()
            deptcode = dr(0)
            areacode = dr(1)
            spos = dr(2)
        Else
            CommUtil.ShowMsg(Me, "沒找到id為" & sid & "之資料,請檢查")
            'Exit Function
        End If
        dr.Close()
        conn.Close()
        SqlCmd = "select deptdesc from dbo.[dept] where deptcode='" & deptcode & "'"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        If (dr.HasRows) Then
            dr.Read()
            dept = dr(0)
        Else
            CommUtil.ShowMsg(Me, "沒找到部門(群組)code為" & deptcode & "之資料,請檢查")
            'Exit Function
        End If
        dr.Close()
        conn.Close()
        SqlCmd = "select areadesc from dbo.[branch] where areacode='" & areacode & "'"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        If (dr.HasRows) Then
            dr.Read()
            area = dr(0)
        Else
            CommUtil.ShowMsg(Me, "沒找到區域code為" & areacode & "之資料,請檢查")
            'Exit Function
        End If
        dr.Close()
        conn.Close()

        SqlCmd = "insert into [dbo].[@XASCH] (docdate,sid,sname,sfid,dept,area,receivedate,spos) " &
                "values(" & "'" & Now_time & "','" & sid & "','" & s_name & "'," & sfid & ",'" & dept & "','" & area & "','" & Now_time & "','" & spos & "')"
        CommUtil.SqlSapExecute("ins", SqlCmd, connsap)
        connsap.Close()

        SqlCmd = "Select max(docnum) from [dbo].[@XASCH]"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        dr.Read()
        num = dr(0)
        dr.Close()
        connsap.Close()
        Return num
    End Function

    Sub InitAddCLHead(Now_time As Date)
        Dim cno, prefix As String
        Dim issuedunit As Integer
        If (Session("grp") = "KS") Then
            prefix = "KS"
            issuedunit = 2
        ElseIf (Session("grp") = "SZ") Then
            prefix = "SZ"
            issuedunit = 3
        Else
            prefix = "TW"
            issuedunit = 1
        End If
        cno = GetCno(Now_time, prefix)

        SqlCmd = "insert into [dbo].[@XSHWS] (docentry,cno,issuedunit) " &
                "values(" & docnum & ",'" & cno & "'," & issuedunit & ")"
        CommUtil.SqlSapExecute("ins", SqlCmd, connsap)
        connsap.Close()
    End Sub
    Sub PutDataToHead()
        Dim status, docdate, cno, issuedperson, subject, content, receiveddpt As String
        Dim issuedunit, issuedtype, receivedunit, sparetype As Integer
        Dim i As Integer
        cno = ""
        subject = ""
        content = ""
        receiveddpt = ""
        docdate = ""
        issuedperson = ""
        SqlCmd = "Select status,docdate,sname,subject " &
        "from [dbo].[@XASCH] " &
        "where docnum=" & docnum
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            dr.Read()
            status = dr(0)
            docdate = Format(dr(1), "yyyy/MM/dd")
            issuedperson = dr(2)
            subject = dr(3)
        End If
        dr.Close()
        connsap.Close()

        SqlCmd = "Select issuedunit,cno,issuedtype,receivedunit,receiveddpt, " &
        "sparetype,content " &
        "from [dbo].[@XSHWS] " &
        "where docentry=" & docnum
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            dr.Read()
            cno = dr(1)
            content = dr(6)
            issuedunit = dr(0)
            issuedtype = dr(2)
            receivedunit = dr(3)
            receiveddpt = dr(4)
            sparetype = dr(5)
        End If
        dr.Close()
        connsap.Close()
        HeadT.Rows(1).Cells(1).Text = docdate
        If (issuedunit <> 0) Then
            RBLCDpt.SelectedValue = issuedunit
        End If
        HeadT.Rows(1).Cells(5).Text = cno
        If (issuedtype <> 0) Then
            RBLType.SelectedValue = issuedtype
        End If
        If (receivedunit <> 0) Then
            RBLDpt.SelectedValue = receivedunit
        End If
        'receiveddpt
        HeadT.Rows(6).Cells(1).Text = issuedperson
        If (sfid = 2 Or sfid = 22) Then
            If (sparetype <> 0) Then
                RBLSpareType.SelectedValue = sparetype
            End If
        End If
        TxtSubject.Text = subject
        TxtInfo.Text = content
        For i = 0 To receiveddpt.Length - 1
            If (receiveddpt.Substring(i, 1) = "1") Then
                ChkDTDpt.Items(i).Selected = True
            End If
        Next

    End Sub

    Sub FillHeadData()
        Dim sfname, sname As String
        Dim dept, area, deptcode, areacode, spos As String
        deptcode = ""
        areacode = ""
        area = ""
        dept = ""
        spos = ""
        sfname = ""
        sname = ""
        SqlCmd = "Select sfname " &
        "from [dbo].[@XSFTT] where sfid=" & sfid
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            dr.Read()
            sfname = dr(0)
        End If
        dr.Close()
        connsap.Close()

        SqlCmd = "select grp,branch,position,name from dbo.[user] where id='" & sid & "'"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        If (dr.HasRows) Then
            dr.Read()
            deptcode = dr(0)
            areacode = dr(1)
            spos = dr(2)
            sname = dr(3)
        Else
            CommUtil.ShowMsg(Me, "沒找到id為" & sid & "之資料,請檢查")
            'Exit Function
        End If
        dr.Close()
        conn.Close()
        SqlCmd = "select deptdesc from dbo.[dept] where deptcode='" & deptcode & "'"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        If (dr.HasRows) Then
            dr.Read()
            dept = dr(0)
        Else
            CommUtil.ShowMsg(Me, "沒找到部門(群組)code為" & deptcode & "之資料,請檢查")
            'Exit Function
        End If
        dr.Close()
        conn.Close()
        SqlCmd = "select areadesc from dbo.[branch] where areacode='" & areacode & "'"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        If (dr.HasRows) Then
            dr.Read()
            area = dr(0)
        Else
            CommUtil.ShowMsg(Me, "沒找到區域code為" & areacode & "之資料,請檢查")
            'Exit Function
        End If
        dr.Close()
        conn.Close()
        If (sfid > 50 And sfid < 80) Then

        Else
            TxtSapNO.Text = "NA"
            TxtPrice.Text = "NA"
            DDLDollorUnit.SelectedIndex = 0
        End If
        HeadT.Rows(0).Cells(3).Text = sname
        HeadT.Rows(1).Cells(1).Text = sfname

        HeadT.Rows(0).Cells(5).Text = dept
        HeadT.Rows(0).Cells(7).Text = area
        Dim formname As String
        formname = ""
        If (maindocnum <> 0) Then
            SqlCmd = "select T1.sfname from  [dbo].[@XASCH] T0 INNER JOIN [dbo].[@XSFTT] T1 ON T0.sfid=T1.sfid where T0.docnum=" & maindocnum
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            formname = dr(0)
            dr.Close()
            connsap.Close()
        End If
        If (sfid <> 100 And sfid <> 101) Then
            TxtAttaDoc.Text = "NA"
        ElseIf (sfid = 100) Then
            If (maindocnum <> 0) Then
                TxtAttaDoc.Text = maindocnum
                TxtSubject.Text = "單號 : " & TxtAttaDoc.Text & "(" & formname & ") 之補充說明事宜"
            End If
        ElseIf (sfid = 101) Then
            If (maindocnum <> 0) Then
                TxtAttaDoc.Text = maindocnum
                TxtSubject.Text = "單號 : " & TxtAttaDoc.Text & "(" & formname & ") 之料件返還事宜"
            End If
        End If
        FileUpLoadObjectDisabled()
    End Sub
    Sub FileUpLoadObjectDisabled()
        CType(FT_m.FindControl("fileul_m"), FileUpload).Enabled = False
        CType(FT_0.FindControl("fileul_0"), FileUpload).Enabled = False
        CType(FT_1.FindControl("fileul_1"), FileUpload).Enabled = False

        CType(FT_m.FindControl("chk_del_m"), CheckBox).Enabled = False
        CType(FT_0.FindControl("chk_del_0"), CheckBox).Enabled = False
        CType(FT_1.FindControl("chk_del_1"), CheckBox).Enabled = False

        CType(FT_m.FindControl("btn_fileact_m"), Button).Enabled = False
        CType(FT_0.FindControl("btn_fileact_0"), Button).Enabled = False
        CType(FT_1.FindControl("btn_fileact_1"), Button).Enabled = False

        'CType(FT_m.FindControl("chk_subfile"), CheckBox).Enabled = False
    End Sub
    Sub FileUpLoadObjectEnabled()
        CType(FT_m.FindControl("fileul_m"), FileUpload).Enabled = True
        CType(FT_0.FindControl("fileul_0"), FileUpload).Enabled = True
        CType(FT_1.FindControl("fileul_1"), FileUpload).Enabled = True

        CType(FT_m.FindControl("chk_del_m"), CheckBox).Enabled = True
        CType(FT_0.FindControl("chk_del_0"), CheckBox).Enabled = True
        CType(FT_1.FindControl("chk_del_1"), CheckBox).Enabled = True

        CType(FT_m.FindControl("btn_fileact_m"), Button).Enabled = True
        CType(FT_0.FindControl("btn_fileact_0"), Button).Enabled = True
        CType(FT_1.FindControl("btn_fileact_1"), Button).Enabled = True

        'CType(FT_m.FindControl("chk_subfile"), CheckBox).Enabled = True
    End Sub
    Sub PutDataToFormInfo()
        Dim status, issuedperson, subject, sfname, priceunit, dept, area, attadoc As String
        Dim sapno As Long
        Dim price As Double
        subject = ""
        sfname = ""
        issuedperson = ""
        sapno = 0
        price = 0
        priceunit = ""
        dept = ""
        area = ""
        SqlCmd = "Select status,docdate,sname,subject,sapno,price,T1.sfname,priceunit,dept,area,attadoc " &
        "from [dbo].[@XASCH] T0 Inner Join [dbo].[@XSFTT] T1 ON T0.sfid=T1.sfid " &
        "where docnum=" & docnum
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            dr.Read()
            status = dr(0)
            'docdate = Format(dr(1), "yyyy/MM/dd")
            issuedperson = dr(2)
            If (Request.QueryString("subject") = "") Then
                subject = dr(3)
            Else
                subject = Request.QueryString("subject")
            End If
            sfname = dr(6)
            sapno = dr(4)
            price = dr(5)
            priceunit = dr(7)
            dept = dr(8)
            area = dr(9)
            'If (maindocnum <> 0) Then
            '    attadoc = maindocnum
            '    TxtAttaDoc.Enabled = False
            '    If (dr(10) <> "NA") Then
            '        If (CLng(dr(10)) <> maindocnum) Then
            '            CommUtil.ShowMsg(Me, "母單欄位與資料庫不同")
            '            Exit Sub
            '        End If
            '    End If
            'Else
            attadoc = dr(10)
            'End If
            HeadT.Rows(0).Cells(1).Text = docnum
            HeadT.Rows(0).Cells(3).Text = issuedperson
            HeadT.Rows(1).Cells(1).Text = sfname
            'If (attadoc <> 0) Then
            TxtAttaDoc.Text = attadoc
            'Else
            'TxtAttaDoc.Text = "NA"
            'End If
            If (sfid > 50 And sfid < 80) Then
                If (act = "fileupanddel" Or act = "fileup" Or act = "filenoassign" Or act = "filedel") Then
                    If (sfid <> 51 And sfid <> 50 And sfid <> 49) Then 'sfid process
                        TxtSapNO.Text = Request.QueryString("sapno")
                        If (Request.QueryString("price") <> 0) Then
                            TxtPrice.Text = Request.QueryString("price") ' Format(Request.QueryString("price"), "###,###.##")
                        Else
                            TxtPrice.Text = 0
                        End If
                        DDLDollorUnit.SelectedIndex = Request.QueryString("unitindex")
                    End If
                    TxtSubject.Text = Request.QueryString("subject")
                Else
                    If (sfid <> 51 And sfid <> 50 And sfid <> 49) Then 'sfid process '把sfid=51 特例排外是因這些欄位是在別處處理(由sql database,若有在add material則在新增時同時更新)
                        If (sapno <> 0) Then
                            TxtSapNO.Text = CStr(sapno)
                        Else
                            TxtSapNO.Text = "NA"
                        End If
                        If (price <> 0) Then
                            TxtPrice.Text = Format(price, "###,###.##")
                        Else
                            TxtPrice.Text = 0
                        End If
                        DDLDollorUnit.SelectedValue = priceunit
                    Else
                        TxtSapNO.Text = "NA"
                    End If
                    TxtSubject.Text = subject
                End If
            Else
                TxtSapNO.Text = "NA"
                TxtPrice.Text = "NA"
                DDLDollorUnit.SelectedIndex = 0
                If (act = "fileupanddel" Or act = "fileup" Or act = "filenoassign" Or act = "filedel") Then ' And (docstatus = "E" Or docstatus = "A")) Then
                    TxtSubject.Text = Request.QueryString("subject")
                Else
                    TxtSubject.Text = subject
                End If
            End If
            HeadT.Rows(0).Cells(5).Text = dept
            HeadT.Rows(0).Cells(7).Text = area

        Else
            HeadT.Rows(0).Cells(1).Text = docnum
            BtnSave.Visible = False
            BtnSend.Visible = False
            BtnDel.Visible = False
            CommUtil.ShowMsg(Me, "沒找到表單號:" & docnum & "之資料")
        End If
        dr.Close()
        connsap.Close()

        '以下可刪除, 只是作驗證用
        'SqlCmd = "Select docdate,sname,subject,sapno,price,T1.sfname,priceunit,dept,area,T2.signdate,T2.uname " &
        '    "from [dbo].[@XASCH] T0 Inner Join [dbo].[@XSFTT] T1 ON T0.sfid=T1.sfid Inner Join [dbo].[@XSPWT] T2 On T0.docnum=T2.docentry " &
        '    "where T0.docnum=" & docnum & " and T2.signprop=0 order by T2.seq"
        'dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        'If (dr.HasRows) Then
        '    Do While (dr.Read())
        '        MsgBox(dr(9) & "--" & dr(5) & "---" & dr(10))
        '    Loop
        'End If
        'dr.Close()
        'connsap.Close()
    End Sub
    Sub CreateCommentField()
        Dim tCell As TableCell
        Dim tRow As TableRow
        CommT.Font.Name = "標楷體"
        tRow = New TableRow()
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.Wrap = False
        tCell.Width = 40
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "簽核意見"
        tRow.Cells.Add(tCell)

        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.Wrap = False
        tCell.HorizontalAlign = HorizontalAlign.Left
        TxtComm = New TextBox
        TxtComm.ID = "txt_comment"
        TxtComm.TextMode = TextBoxMode.MultiLine
        TxtComm.Rows = 4
        TxtComm.Width = 1000
        TxtComm.Font.Name = "標楷體"
        'TxtComm.BackColor = Drawing.Color.Cornsilk
        tCell.Controls.Add(TxtComm)
        tRow.Cells.Add(tCell)
        CommT.Rows.Add(tRow)
        If (docstatus <> "A" And docstatus <> "E") Then
            CommT.Enabled = True
        Else
            CommT.Enabled = False
            TxtComm.Text = "此處出現送審時才能填寫"
        End If
    End Sub
    Sub CreateSignFlowHistoryField()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim width As Integer
        Dim i As Integer
        Dim j As Integer
        Dim maxseq As Integer
        Dim uid, uname As String
        Dim keyuid, keyuname As String
        Dim keyseq, signprop, nextsignprop As Integer
        Dim formstatus As String
        formstatus = ""
        i = 0
        width = 120
        SqlCmd = "select status from  [dbo].[@XASCH] where docnum=" & docnum
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            dr.Read()
            formstatus = dr(0)
        End If
        dr.Close()
        connsap.Close()
        SqlCmd = "select uid,uname,seq,signprop from [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum & " and status=1"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        keyuid = ""
        keyuname = ""
        keyseq = 0
        If (dr.HasRows) Then
            dr.Read()
            keyuid = dr(0)
            keyuname = dr(1)
            keyseq = dr(2)
            signprop = dr(3)
        End If
        dr.Close()
        connsap.Close()
        If (signprop = 0) Then
            SqlCmd = "select signprop from [dbo].[@XSPWT] where docentry=" & docnum & " and seq=" & keyseq + 1
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                dr.Read()
                nextsignprop = dr(0)
            End If
            dr.Close()
            connsap.Close()
        End If

        SqlCmd = "select uname,status,comment,signdate,uid,agnname from  [dbo].[@XSPHT] where docentry=" & docnum & " order by flowseq"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            tRow = New TableRow()
            'If (i Mod 2) Then
            'tRow.BackColor = Drawing.Color.Cornsilk
            'End If
            tCell = New TableCell()
            'tCell.BorderWidth = 1
            tCell.Width = width
            tCell.Wrap = False
            tCell.Text = "關卡"
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Cells.Add(tCell)

            tCell = New TableCell() '動作
            'tCell.BorderWidth = 1
            tCell.Width = width
            tCell.Wrap = False
            tCell.Text = "動作"
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Cells.Add(tCell)

            tCell = New TableCell() '簽核意見
            'tCell.BorderWidth = 1
            tCell.ColumnSpan = 5
            tCell.Wrap = False
            tCell.Width = width
            tCell.Text = "簽核意見"
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Cells.Add(tCell)

            tCell = New TableCell() '代理人
            'tCell.BorderWidth = 1
            tCell.Width = width
            tCell.Wrap = False
            tCell.Text = "代理人啟用"
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Cells.Add(tCell)

            tCell = New TableCell() '日期
            'tCell.BorderWidth = 1
            tCell.Width = width
            tCell.Wrap = False
            tCell.ColumnSpan = 2
            tCell.Text = "日期"
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Cells.Add(tCell)
            HT.Rows.Add(tRow)
            Do While (dr.Read())
                i = i + 1
                tRow = New TableRow()
                'If (i Mod 2) Then
                tRow.BackColor = Drawing.Color.Cornsilk
                'End If
                tCell = New TableCell()
                'tCell.BorderWidth = 1
                tCell.BackColor = Drawing.Color.Cornsilk
                tCell.Width = width
                tCell.Wrap = False
                tCell.Text = dr(0)
                tCell.HorizontalAlign = HorizontalAlign.Center
                tRow.Cells.Add(tCell)

                tCell = New TableCell() '動作
                'tCell.BorderWidth = 1
                If (dr(1) <> "反對") Then
                    tCell.BackColor = Drawing.Color.Cornsilk
                Else
                    tCell.BackColor = Drawing.Color.Red
                End If
                tCell.Width = width
                tCell.Wrap = False
                tCell.Text = dr(1)
                tCell.HorizontalAlign = HorizontalAlign.Center
                tRow.Cells.Add(tCell)

                tCell = New TableCell() '簽核意見
                'tCell.BorderWidth = 1
                tCell.BackColor = Drawing.Color.Cornsilk
                tCell.ColumnSpan = 5
                tCell.Wrap = False
                tCell.Width = width
                tCell.Text = dr(2)
                tCell.ToolTip = dr(2)
                If (tCell.Text.Length > 80) Then
                    tCell.Text = tCell.Text.Substring(0, 80) + "..."
                End If
                tCell.HorizontalAlign = HorizontalAlign.Left
                tRow.Cells.Add(tCell)

                tCell = New TableCell() '代理人
                'tCell.BorderWidth = 1
                tCell.BackColor = Drawing.Color.Cornsilk
                tCell.Width = width
                tCell.Wrap = False
                tCell.Text = dr(5)
                tCell.HorizontalAlign = HorizontalAlign.Center
                tRow.Cells.Add(tCell)

                tCell = New TableCell() '日期
                'tCell.BorderWidth = 1
                tCell.BackColor = Drawing.Color.Cornsilk
                tCell.Width = width
                tCell.Wrap = False
                tCell.ColumnSpan = 2
                tCell.Text = dr(3)
                tCell.HorizontalAlign = HorizontalAlign.Center
                tRow.Cells.Add(tCell)
                HT.Rows.Add(tRow)
            Loop
            tRow = New TableRow()
            tRow.BackColor = Drawing.Color.Cornsilk
            SqlCmd = "select IsNull(max(seq),1) from [dbo].[@XSPWT] where signprop= 0 And docentry = " & docnum
            dr1 = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap1)
            If (dr1.HasRows) Then
                dr1.Read()
                maxseq = dr1(0)
                dr1.Close()
                connsap1.Close()
            End If
            If (formstatus = "F" Or formstatus = "C" Or formstatus = "T") Then '最後一核決者或表單已停止
                For j = 0 To 4
                    tCell = New TableCell()
                    tCell.Width = width
                    tCell.BackColor = Drawing.Color.Cornsilk
                    tCell.Wrap = False
                    If (j = 1) Then
                        tCell.Text = "流程結束"
                    ElseIf (j = 2) Then
                        tCell.ColumnSpan = 5
                    ElseIf (j = 4) Then
                        tCell.ColumnSpan = 2
                    End If
                    tCell.HorizontalAlign = HorizontalAlign.Center
                    tRow.Cells.Add(tCell)
                Next
                HT.Rows.Add(tRow)
            Else
                tRow = New TableRow()
                tRow.BackColor = Drawing.Color.Cornsilk
                'SqlCmd = "select uname,uid from [dbo].[@XSPWT] where docentry=" & docnum & " and seq=" & keyseq
                'dr1 = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap1)
                'dr1.Read()
                'uname = dr1(0)
                'uid = dr1(1)
                'dr1.Close()
                'connsap1.Close()
                ''MsgBox(i & " " & uid)
                ''If (dr1(1) <> Session("s_id")) Then
                For j = 0 To 4
                    tCell = New TableCell()
                    tCell.BackColor = Drawing.Color.Cornsilk
                    tCell.Width = width
                    tCell.Wrap = False
                    If (j = 0) Then
                        tCell.Text = keyuname
                    ElseIf (j = 1) Then
                        If (formstatus = "B") Then
                            tCell.Text = "重啟流程"
                        Else
                            If (keyuid <> Session("s_id")) Then
                                tCell.Text = "目前關卡"
                            Else
                                tCell.Text = "關卡覆核中"
                            End If
                        End If
                    ElseIf (j = 2) Then
                        tCell.ColumnSpan = 5
                    ElseIf (j = 3) Then
                        If (keyuid = Session("s_id")) Then
                            Dim connL As New SqlConnection
                            Dim drL As SqlDataReader
                            Dim agnname As String
                            agnname = ""
                            If (agnidG <> "") Then
                                SqlCmd = "select name from dbo.[user] where id='" & agnidG & "'"
                                drL = CommUtil.SelectLocalSqlUsingDr(SqlCmd, connL)
                                If (drL.HasRows) Then
                                    drL.Read()
                                    agnname = drL(0)
                                End If
                                drL.Close()
                                connL.Close()
                            End If
                            tCell.Text = agnname
                        End If
                    ElseIf (j = 4) Then
                        tCell.ColumnSpan = 2
                    End If
                    tCell.HorizontalAlign = HorizontalAlign.Center
                    tRow.Cells.Add(tCell)
                Next
                HT.Rows.Add(tRow)
                If (keyuid = Session("s_id") And formstatus <> "B") Then
                    tRow = New TableRow()
                    tRow.BackColor = Drawing.Color.Cornsilk
                    If (maxseq = keyseq) Then '歷史簽核下一筆是最後一筆==> 產生無下一關
                        For j = 0 To 4
                            tCell = New TableCell()
                            tCell.Width = width
                            tCell.BackColor = Drawing.Color.Cornsilk
                            tCell.Wrap = False
                            If (j = 1) Then
                                tCell.Text = "無下一關"
                            ElseIf (j = 2) Then
                                tCell.ColumnSpan = 5
                            ElseIf (j = 4) Then
                                tCell.ColumnSpan = 2
                            End If
                            tCell.HorizontalAlign = HorizontalAlign.Center
                            tRow.Cells.Add(tCell)
                        Next
                    Else '歷史簽核下一筆不是最後一筆===>產生下一關卡
                        SqlCmd = "select uname,uid from [dbo].[@XSPWT] where docentry=" & docnum & " and seq=" & keyseq + 1
                        dr1 = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap1)
                        dr1.Read()
                        uname = dr1(0)
                        uid = dr1(1)
                        dr1.Close()
                        connsap1.Close()
                        'If (dr1(1) <> Session("s_id")) Then
                        For j = 0 To 4
                            tCell = New TableCell()
                            tCell.BackColor = Drawing.Color.Cornsilk
                            tCell.Width = width
                            tCell.Wrap = False
                            If (j = 0) Then
                                tCell.Text = uname
                            ElseIf (j = 1) Then
                                tCell.Text = "下一關"
                            ElseIf (j = 2) Then
                                tCell.ColumnSpan = 5
                            ElseIf (j = 4) Then
                                tCell.ColumnSpan = 2
                            End If
                            tCell.HorizontalAlign = HorizontalAlign.Center
                            tRow.Cells.Add(tCell)
                        Next
                    End If
                    HT.Rows.Add(tRow)
                End If
            End If
        End If
        dr.Close()
        connsap.Close()
    End Sub
    Sub CreateFormInfo()
        TxtPrice = New TextBox
        'TxtPrice.Width = 80
        HeadT.Font.Name = "標楷體"
        HeadT.Font.Size = 12
        HeadT.Rows(1).Cells(5).Controls.Add(TxtPrice)
        DDLDollorUnit = New DropDownList
        DDLDollorUnit.Items.Add("選擇幣別")
        DDLDollorUnit.Items.Add("NTD")
        DDLDollorUnit.Items.Add("USD")
        DDLDollorUnit.Items.Add("RMB")
        DDLDollorUnit.SelectedIndex = 0
        HeadT.Rows(1).Cells(6).Controls.Add(DDLDollorUnit)
        TxtSubject = New TextBox
        TxtSubject.Width = 1000
        TxtSubject.Font.Name = "標楷體"
        HeadT.Rows(2).Cells(1).Controls.Add(TxtSubject)

        TxtAttaDoc = New TextBox
        'TxtAttaDoc.Width =80
        TxtAttaDoc.Font.Name = "標楷體"
        HeadT.Rows(2).Cells(3).Controls.Add(TxtAttaDoc)
        TxtAttaDoc.AutoPostBack = True
        AddHandler TxtAttaDoc.TextChanged, AddressOf TxtAttaDoc_TextChanged

        TxtSapNO = New TextBox
        If (sfid > 70 And sfid < 80) Then
            TxtSapNO.AutoPostBack = True
            AddHandler TxtSapNO.TextChanged, AddressOf TxtSapNO_TextChanged
        End If
        HeadT.Rows(1).Cells(3).Controls.Add(TxtSapNO)
        If (sfid > 50 And sfid < 80) Then
            If (sfid > 70 And sfid < 80) Then
                TxtSapNO.Enabled = True
            Else
                TxtSapNO.Enabled = False
            End If
            TxtPrice.Enabled = True
            DDLDollorUnit.Enabled = True
            If (sfid > 70 And sfid < 80) Then
                TxtPrice.Enabled = False
                DDLDollorUnit.Enabled = False
            End If
        Else
            TxtSapNO.Enabled = False
            TxtPrice.Enabled = False
            DDLDollorUnit.Enabled = False
        End If
        If (docstatus = "E" Or docstatus = "D" Or docstatus = "B" Or docstatus = "R") Then
            TxtSubject.Enabled = True
            TxtAttaDoc.Enabled = True
        Else
            TxtSubject.Enabled = False
            TxtSapNO.Enabled = False
            TxtPrice.Enabled = False
            DDLDollorUnit.Enabled = False
            TxtAttaDoc.Enabled = False
        End If
        If (sfid = 100) Then
            HeadT.Rows(2).Cells(2).Text = "被補充單號"
            TxtSubject.Enabled = False
        ElseIf (sfid = 101) Then
            HeadT.Rows(2).Cells(2).Text = "離倉/借入單號"
            TxtSubject.Enabled = False
        ElseIf (sfid = 23 Or sfid = 24) Then
            HeadT.Rows(2).Cells(2).Text = "返還單號"
            TxtSubject.Enabled = False
        Else
            TxtAttaDoc.Enabled = False
            HeadT.Rows(2).Cells(2).Text = "補充單號"
        End If
        'HeadT 欄位已由aspx規劃完成
    End Sub
    '    Sub TxtReason_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
    'SqlCmd = "update [dbo].[@XSMLS] set descrip='" & TxtReason.Text & "'" &
    '                " where docentry=" & docnum & " and head=1"
    'CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
    'connsap.Close()
    'CommUtil.ShowMsg(Me, "事由有更動並已存檔")

    '    End Sub
    Sub TxtSapNO_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
        If (sfid > 70 And sfid < 80) Then '各式採購單
            '先check 是否之前已有送審過
            SqlCmd = "Select count(*) from [dbo].[@XASCH] " &
                    "where sapno=" & CLng(TxtSapNO.Text) & " and status<>'C'"
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            If (dr(0) <> 0) Then
                CommUtil.ShowMsg(Me, "此PO單之前已建立過 , 請確認此為更新")
            End If
            dr.Close()
            connsap.Close()

            SqlCmd = "select T0.DocCur,T0.DocTotalSy,T0.DocTotalFC,IsNull(T0.Comments,'') from dbo.OPOR T0 where T0.docnum=" & TxtSapNO.Text
            'MsgBox(SqlCmd)
            'Exit Sub
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                dr.Read()
                If (dr(0) = "NTD") Then
                    TxtPrice.Text = Format(dr(1), "###,###.##")
                Else
                    TxtPrice.Text = Format(dr(2), "###,###.##")
                End If
                DDLDollorUnit.SelectedValue = dr(0)
                'If (TxtSubject.Text = "") Then
                If (dr(3) <> "") Then
                    TxtSubject.Text = "採購單:" & TxtSapNO.Text & "---" & dr(3)
                Else
                    TxtSubject.Text = "採購單:" & TxtSapNO.Text
                End If
                'End If
            Else
                CommUtil.ShowMsg(Me, "在Sap找不到單號:" & TxtSapNO.Text & "之採購單")
            End If
            dr.Close()
            connsap.Close()
        End If
    End Sub
    Sub tTxtreturn_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim Txtx As TextBox = sender
        Dim inum, outqty, rtnqty, nowrtnqty As Long
        Dim trow As Integer
        Dim str() As String
        str = Split(Txtx.ID, "_")
        trow = str(1)
        inum = str(2)
        If (IsNumeric(Txtx.Text)) Then
            outqty = CLng(ContentT.Rows(trow).Cells(3).Text)
            rtnqty = CLng(ContentT.Rows(trow).Cells(4).Text)
            nowrtnqty = CLng(Txtx.Text)
            If ((nowrtnqty + rtnqty) >= 0) Then
                If (nowrtnqty <= (outqty - rtnqty)) Then
                    SqlCmd = "update dbo.[@XSMLS] set nowrtnqty= " & nowrtnqty & " where num=" & inum
                    CommUtil.SqlSapExecute("upd", SqlCmd, conn)
                    conn.Close()
                Else
                    CommUtil.ShowMsg(Me, "可還" & outqty - rtnqty & "個 , 但輸入之返還數量(" & nowrtnqty & ")已超過")
                    Txtx.Text = 0
                End If
                If (nowrtnqty < 0) Then
                    CommUtil.ShowMsg(Me, "!!!! 請注意你輸入的數字為 '負數'")
                End If
            Else
                Txtx.Text = 0
                CommUtil.ShowMsg(Me, "注意你輸入的數字為 '負數', 但回沖的數量已大於已返還數量, 請輸入正確數量")
            End If
        Else
            Txtx.Text = 0
            CommUtil.ShowMsg(Me, "需為數字")
        End If
    End Sub
    Sub tTxtNote_TextChanged(ByVal sender As Object, ByVal e As EventArgs) 'qqqqq
        Dim Txtx As TextBox = sender
        Dim inum As Long
        Dim str() As String
        str = Split(Txtx.ID, "_")
        inum = str(2)
        SqlCmd = "update dbo.[@XSMLS] set comment='" & Txtx.Text & "' where num=" & inum
        CommUtil.SqlSapExecute("upd", SqlCmd, conn)
        conn.Close()
    End Sub
    Sub TxtAttaDoc_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
        If (Not IsNumeric(TxtAttaDoc.Text)) Then
            CommUtil.ShowMsg(Me, "單號需為整數字")
            TxtAttaDoc.Text = PreAttaDoc
            TxtSubject.Text = ""
            Exit Sub
        End If
        If (sfid = 101) Then
            'Check是否為離倉單
            SqlCmd = "Select mtype FROM [dbo].[@XSMLS] T0 WHERE T0.[docentry] =" & CLng(TxtAttaDoc.Text) & " and head=1"
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, conn)
            If (dr.HasRows) Then
                dr.Read()
                If (dr(0) = 0) Then
                    CommUtil.ShowMsg(Me, "輸入之母單號不是離倉(借入)單號")
                    dr.Close()
                    conn.Close()
                    TxtAttaDoc.Text = PreAttaDoc
                    TxtSubject.Text = ""
                    Exit Sub
                End If
            Else
                CommUtil.ShowMsg(Me, "無此母單號可返還")
                dr.Close()
                conn.Close()
                TxtAttaDoc.Text = PreAttaDoc
                TxtSubject.Text = ""
                Exit Sub
            End If
            dr.Close()
            conn.Close()
            '若選擇之母單號有另外單號在簽核 , 要等其簽完
            'SqlCmd = "Select sum(nowrtnqty) FROM [dbo].[@XSMLS] T0 WHERE T0.[docentry] =" & CLng(TxtAttaDoc.Text) & " and head=0"
            SqlCmd = "Select docnum FROM [dbo].[@XASCH] T0 WHERE T0.[attadoc] ='" & TxtAttaDoc.Text & "' and sfid=101 and status<>'F' and status<>'T' and status<>'C'"
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, conn)
            If (dr.HasRows) Then
                dr.Read()
                'If (dr(0) <> 0) Then
                CommUtil.ShowMsg(Me, "目前還有以母單號:" & TxtAttaDoc.Text & " 為返還單之單據(" & dr(0) & ")在簽核,需等其簽完才能新增")
                dr.Close()
                conn.Close()
                TxtAttaDoc.Text = PreAttaDoc
                TxtSubject.Text = ""
                Exit Sub
                'End If
            End If
            dr.Close()
            conn.Close()
            'Check 此單是否還有料件要返還
            SqlCmd = "Select sum(quantity),sum(rtnqty) FROM [dbo].[@XSMLS] T0 WHERE T0.[docentry] =" & CLng(TxtAttaDoc.Text) & " and head=0"
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, conn)
            dr.Read()
            If (dr(0) = dr(1)) Then
                CommUtil.ShowMsg(Me, "此母單號:" & TxtAttaDoc.Text & " 之料件已全部返還,請改輸入其他母單號")
                dr.Close()
                conn.Close()
                TxtAttaDoc.Text = PreAttaDoc
                TxtSubject.Text = ""
                Exit Sub
            End If
            dr.Close()
            conn.Close()
        End If
        TxtSubject.Text = ""
        'If (IsNumeric(TxtAttaDoc.Text)) Then
        SqlCmd = "Select status from [dbo].[@XASCH] " &
            "where docnum=" & CLng(TxtAttaDoc.Text)
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            dr.Read()
            If (dr(0) <> "F" And dr(0) <> "T") Then
                CommUtil.ShowMsg(Me, "此(母)單號:" & TxtAttaDoc.Text & " 尚未完成簽核, 請等其簽核完畢")
                TxtAttaDoc.Text = PreAttaDoc
            End If
        Else
            CommUtil.ShowMsg(Me, "無此被附屬(母)單號:" & TxtAttaDoc.Text)
            TxtAttaDoc.Text = PreAttaDoc
        End If
        dr.Close()
        connsap.Close()
        'Else
        '    If (TxtAttaDoc.Text <> "") Then
        '        CommUtil.ShowMsg(Me, "單號需為整數字")
        '        TxtAttaDoc.Text = ""
        '    End If
        'End If
        If (TxtAttaDoc.Text <> "") Then
            Dim formname As String
            formname = ""
            If (maindocnum <> 0) Then
                SqlCmd = "select T1.sfname from  [dbo].[@XASCH] T0 INNER JOIN [dbo].[@XSFTT] T1 ON T0.sfid=T1.sfid where T0.docnum=" & CLng(TxtAttaDoc.Text)
                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                dr.Read()
                formname = dr(0)
                dr.Close()
                connsap.Close()
            End If
            If (sfid = 100) Then
                TxtSubject.Text = "單號 : " & TxtAttaDoc.Text & "(" & formname & ") 之補充說明事宜"
            ElseIf (sfid = 101) Then
                TxtSubject.Text = "單號 : " & TxtAttaDoc.Text & "(" & formname & ") 之料件返還事宜"
                SqlCmd = "update dbo.[@XSMLS] set nowrtnqty=0 where docentry=" & CLng(PreAttaDoc)
                CommUtil.SqlSapExecute("upd", SqlCmd, conn)
                conn.Close()
            End If
        End If
        SqlCmd = "update [dbo].[@XASCH] set attadoc='" & TxtAttaDoc.Text & "',subject='" & TxtSubject.Text & "' where docnum=" & docnum
        CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
        connsap.Close()
        If (sfid = 101) Then
            If (maindocnum <> 0) Then
                Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&actmode=single&act=add101&sfid=101" &
                    "&formtypeindex=" & formtypeindex & "&inchargeindex=" & inchargeindex &
                    "&traceindex=" & traceindex & "&maindocnum=" & maindocnum & "&docnum=" & docnum)
            Else
                Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&status=" & docstatus &
                                "&docnum=" & docnum &
                                "&actmode=single&formstatusindex=" & formstatusindex &
                                "&formtypeindex=" & formtypeindex & "&sfid=" & sfid & "&signflowmode=" & Request.QueryString("signflowmode"))
            End If
        Else
            GenAttaFileList()
            ShowIframeContent()
        End If
    End Sub
    Sub CreateSignOffFlowField(i As Integer)
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim TxtComm As TextBox
        Dim tImage As Image
        Dim width As Integer
        If (i Mod 2) Then
            SignT.Font.Name = "標楷體"
            SignT.Font.Size = 12
            width = 80
            tRow = New TableRow()
            'If (i Mod 2) Then
            'tRow.BackColor = Drawing.Color.Cornsilk
            'End If
            tCell = New TableCell() '職位資訊_L
            tCell.BorderWidth = 1
            tCell.Width = width
            tCell.Wrap = False
            tCell.ColumnSpan = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Cells.Add(tCell)

            tCell = New TableCell() 'image_L
            tCell.BorderWidth = 1
            tCell.Width = width
            tCell.RowSpan = 3
            tCell.Wrap = False
            tCell.HorizontalAlign = HorizontalAlign.Center
            tImage = New Image
            tImage.ID = "image_signL_" & i
            tCell.Controls.Add(tImage)
            tRow.Cells.Add(tCell)

            tCell = New TableCell() '意見_L
            tCell.BorderWidth = 1
            tCell.RowSpan = 3
            tCell.ColumnSpan = 2
            tCell.Wrap = True
            tCell.Width = width
            tCell.HorizontalAlign = HorizontalAlign.Left
            TxtComm = New TextBox
            TxtComm.ID = "txt_commL_" & i
            TxtComm.TextMode = TextBoxMode.MultiLine
            TxtComm.Rows = 4
            TxtComm.Width = 380 'pixel
            TxtComm.Font.Size = 12 'point
            TxtComm.BorderWidth = 0
            TxtComm.Enabled = False
            tCell.Controls.Add(TxtComm)
            tRow.Cells.Add(tCell)

            tCell = New TableCell() '職位資訊_R '因此列此Cell為此列第四個create , 故序號為3
            tCell.BorderWidth = 1
            tCell.Width = width
            tCell.Wrap = False
            tCell.BackColor = Drawing.Color.Cornsilk
            tCell.ColumnSpan = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Cells.Add(tCell)

            tCell = New TableCell() 'image_R
            tCell.BorderWidth = 1
            tCell.Width = width
            tCell.RowSpan = 3
            tCell.Wrap = False
            tCell.BackColor = Drawing.Color.Cornsilk
            tCell.HorizontalAlign = HorizontalAlign.Center
            tImage = New Image
            tImage.ID = "image_signR_" & i + 1
            tCell.Controls.Add(tImage)
            tRow.Cells.Add(tCell)

            tCell = New TableCell() '意見_R
            tCell.BorderWidth = 1
            tCell.RowSpan = 3
            tCell.ColumnSpan = 2
            tCell.BackColor = Drawing.Color.Cornsilk
            tCell.Wrap = True
            tCell.Width = width
            tCell.HorizontalAlign = HorizontalAlign.Left
            TxtComm = New TextBox
            TxtComm.ID = "txt_commR_" & i + 1
            TxtComm.TextMode = TextBoxMode.MultiLine
            TxtComm.Rows = 4
            TxtComm.Width = 380 'pixel
            TxtComm.Font.Size = 12 'point
            TxtComm.BorderWidth = 0
            TxtComm.Enabled = False
            TxtComm.BackColor = Drawing.Color.Cornsilk
            tCell.Controls.Add(TxtComm)
            tRow.Cells.Add(tCell)
            SignT.Rows.Add(tRow)

            'row=1
            tRow = New TableRow()
            tCell = New TableCell() '簽核人資訊L'因此列此Cell為此列第一個create , 故序號為0
            tCell.BorderWidth = 1
            tCell.Width = width
            tCell.Wrap = False
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Cells.Add(tCell)

            tCell = New TableCell() '簽核人資訊R '因此列此Cell為此列第二個create , 故序號為1
            tCell.BorderWidth = 1
            tCell.Width = width
            tCell.Wrap = False
            tCell.BackColor = Drawing.Color.Cornsilk
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Cells.Add(tCell)
            SignT.Rows.Add(tRow)
            'row=2
            tRow = New TableRow()
            tCell = New TableCell() '簽核日期資訊L
            tCell.BorderWidth = 1
            tCell.Width = width
            tCell.Wrap = False
            tCell.ColumnSpan = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Cells.Add(tCell)
            tCell = New TableCell() '簽核日期資訊R
            tCell.BorderWidth = 1
            tCell.Width = width
            tCell.Wrap = False
            tCell.BackColor = Drawing.Color.Cornsilk
            tCell.ColumnSpan = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Cells.Add(tCell)
            SignT.Rows.Add(tRow)
        End If
    End Sub
    Sub PutDataToSignOffFlow()
        Dim uid, uname, upos, comment, deptdesc, areadesc As String
        Dim seq, status, row As Integer
        Dim signdate, agnid As String
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim rowsbychara, rowsbynewline, memorows As Integer
        deptdesc = ""
        areadesc = ""
        row = 0
        'If (docstatus = "F" Or docstatus = "T") Then
        SqlCmd = "select uid,uname,upos,comment,seq,status,IsNull(signdate,''),agnid from [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum & " order by seq"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            Do While (dr.Read())
                uid = dr(0)
                uname = dr(1)
                If (dr(2) <> "NA") Then
                    upos = dr(2)
                Else
                    upos = ""
                End If
                comment = dr(3)
                seq = dr(4)
                status = dr(5)
                agnid = dr(7)
                If (dr(6) = "1900/01/01") Then
                    signdate = "NA"
                Else
                    signdate = dr(6)
                End If
                'row = (seq - 1) * 3
                SqlCmd = "select T1.deptdesc,T2.areadesc,T0.position from dbo.[user] T0 Inner join dbo.[dept] T1 on T0.grp=T1.deptcode " &
                        "Inner Join dbo.[branch] T2 on T0.branch=T2.areacode where T0.id='" & uid & "'"
                dr1 = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                If (dr1.HasRows) Then
                    dr1.Read()
                    deptdesc = dr1(0)
                    areadesc = dr1(1)
                Else
                    CommUtil.ShowMsg(Me, "沒找到id為" & uid & "之資料,請檢查")
                End If
                dr1.Close()
                conn.Close()
                If (seq Mod 2) Then
                    SignT.Rows(row).Cells(0).Text = areadesc & "  " & deptdesc 'upos
                    'SignT.Rows(row + 1).Cells(0).Text = uname
                    SignT.Rows(row + 2).Cells(0).Text = signdate
                    If (agnid = "") Then
                        SignT.Rows(row + 1).Cells(0).Text = uname & " " & upos '"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & uname & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                    Else
                        Dim agnname As String
                        agnname = ""
                        SqlCmd = "select name from dbo.[user] where id='" & agnid & "'"
                        drL = CommUtil.SelectLocalSqlUsingDr(SqlCmd, connL)
                        If (drL.HasRows) Then
                            drL.Read()
                            agnname = drL(0)
                        End If
                        drL.Close()
                        connL.Close()
                        'SignT.Rows(row).Cells(1).Text = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & uname & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>代理人:" & agnname
                        SignT.Rows(row + 1).Cells(0).Text = uname & "&nbsp;&nbsp;代理人:" & agnname
                    End If
                    If (sfid = 16 And seq = 1) Then
                        SqlCmd = "Select T0.id,T0.createid,T0.idname FROM [dbo].[@XRSCT] T0 WHERE T0.[docentry] =" & docnum
                        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                        If (drL.HasRows) Then
                            drL.Read()
                            If (drL(0) <> drL(1)) Then
                                SignT.Rows(row + 1).Cells(0).Text = SignT.Rows(row + 1).Cells(0).Text & "(替 " & drL(2) & " 發送)"
                            End If
                        End If
                        drL.Close()
                        connL.Close()
                    End If

                    rowsbynewline = System.Text.RegularExpressions.Regex.Matches(comment, "\r\n").Count + 1
                    rowsbychara = (comment.Length / 40) + 1
                    If (rowsbynewline >= rowsbychara) Then
                        memorows = rowsbynewline
                    Else
                        memorows = rowsbychara
                    End If
                    If (memorows <= 4) Then
                        CType(SignT.FindControl("txt_commL_" & seq), TextBox).Rows = 4
                    Else
                        CType(SignT.FindControl("txt_commL_" & seq), TextBox).Rows = memorows
                    End If

                    CType(SignT.FindControl("txt_commL_" & seq), TextBox).Text = comment
                    If (status = 2 Or status = 100) Then
                        CType(SignT.FindControl("image_signL_" & seq), Image).ImageUrl = "~/image/ok1.jpg"
                    ElseIf (status = 3 Or status = 5) Then
                        CType(SignT.FindControl("image_signL_" & seq), Image).ImageUrl = "~/image/rj1.jpg"
                    ElseIf (status = 10) Then
                        CType(SignT.FindControl("image_signL_" & seq), Image).ImageUrl = "~/image/skip1.jpg"
                    End If
                Else
                    SignT.Rows(row).Cells(3).Text = areadesc & "  " & deptdesc 'upos '因此列此Cell為此列第四個create , 故序號為3
                    'SignT.Rows(row + 1).Cells(1).Text = uname '因此列此Cell為此列第二個create , 故序號為1
                    SignT.Rows(row + 2).Cells(1).Text = signdate
                    If (agnid = "") Then
                        SignT.Rows(row + 1).Cells(1).Text = uname & " " & upos '"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & uname & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                    Else
                        Dim agnname As String
                        agnname = ""
                        SqlCmd = "select name from dbo.[user] where id='" & agnid & "'"
                        drL = CommUtil.SelectLocalSqlUsingDr(SqlCmd, connL)
                        If (drL.HasRows) Then
                            drL.Read()
                            agnname = drL(0)
                        End If
                        drL.Close()
                        connL.Close()
                        'SignT.Rows(row).Cells(1).Text = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & uname & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>代理人:" & agnname
                        SignT.Rows(row + 1).Cells(1).Text = uname & "&nbsp;&nbsp;待理人:" & agnname
                    End If
                    If (sfid = 16 And seq = 1) Then
                        SqlCmd = "Select T0.id,T0.createid,T0.idname FROM [dbo].[@XRSCT] T0 WHERE T0.[docentry] =" & docnum
                        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                        If (drL.HasRows) Then
                            drL.Read()
                            If (drL(0) <> drL(1)) Then
                                SignT.Rows(row + 1).Cells(1).Text = SignT.Rows(row + 1).Cells(1).Text & "(替 " & drL(2) & " 發送)"
                            End If
                        End If
                        drL.Close()
                        connL.Close()
                    End If
                    'SignT.Rows(row).Cells(5).Text = comment
                    rowsbynewline = System.Text.RegularExpressions.Regex.Matches(comment, "\r\n").Count + 1
                    rowsbychara = (comment.Length / 40) + 1
                    If (rowsbynewline >= rowsbychara) Then
                        memorows = rowsbynewline
                    Else
                        memorows = rowsbychara
                    End If
                    If (memorows <= 4) Then
                        CType(SignT.FindControl("txt_commR_" & seq), TextBox).Rows = 4
                    Else
                        CType(SignT.FindControl("txt_commR_" & seq), TextBox).Rows = memorows
                    End If
                    CType(SignT.FindControl("txt_commR_" & seq), TextBox).Text = comment
                    If (status = 2 Or status = 100) Then
                        CType(SignT.FindControl("image_signR_" & seq), Image).ImageUrl = "~/image/ok1.jpg"
                    ElseIf (status = 3 Or status = 5) Then
                        CType(SignT.FindControl("image_signR_" & seq), Image).ImageUrl = "~/image/rj1.jpg"
                    ElseIf (status = 10) Then
                        CType(SignT.FindControl("image_signR_" & seq), Image).ImageUrl = "~/image/skip1.jpg"
                    End If
                    row = row + 3
                End If
            Loop
        End If
        dr.Close()
        connsap.Close()
        'End If
    End Sub
    Sub PutDataToSignOffFlow_ORG()
        Dim uid, uname, upos, comment As String
        Dim seq, status, row As Integer
        Dim signdate, agnid As String
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        'If (docstatus = "F" Or docstatus = "T") Then
        SqlCmd = "select uid,uname,upos,comment,seq,status,IsNull(signdate,''),agnid from [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum & " order by seq"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            Do While (dr.Read())
                uid = dr(0)
                uname = dr(1)
                upos = dr(2)
                comment = dr(3)
                seq = dr(4)
                status = dr(5)
                agnid = dr(7)
                If (dr(6) = "1900/01/01") Then
                    signdate = "NA"
                Else
                    signdate = dr(6)
                End If
                row = (seq - 1) * 2
                SignT.Rows(row).Cells(0).Text = upos
                'SignT.Rows(row + 1).Cells(0).Text = uname
                SignT.Rows(row + 1).Cells(0).Text = signdate
                If (agnid = "") Then
                    SignT.Rows(row).Cells(1).Text = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & uname & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                Else
                    Dim agnname As String
                    agnname = ""
                    SqlCmd = "select name from dbo.[user] where id='" & agnid & "'"
                    drL = CommUtil.SelectLocalSqlUsingDr(SqlCmd, connL)
                    If (drL.HasRows) Then
                        drL.Read()
                        agnname = drL(0)
                    End If
                    drL.Close()
                    connL.Close()
                    SignT.Rows(row).Cells(1).Text = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & uname & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>代理人:" & agnname
                End If
                If (sfid = 16 And seq = 1) Then
                    SqlCmd = "Select T0.id,T0.createid,T0.idname FROM [dbo].[@XRSCT] T0 WHERE T0.[docentry] =" & docnum
                    drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                    If (drL.HasRows) Then
                        drL.Read()
                        If (drL(0) <> drL(1)) Then
                            SignT.Rows(row).Cells(1).Text = SignT.Rows(row).Cells(1).Text & "<br>(代替 " & drL(2) & " 發送)"
                        End If
                    End If
                    drL.Close()
                    connL.Close()
                End If
                If ((System.Text.RegularExpressions.Regex.Matches(comment, "\r\n").Count + 1) <= 4) Then
                    CType(SignT.FindControl("txt_comm_" & seq), TextBox).Rows = 4
                Else
                    CType(SignT.FindControl("txt_comm_" & seq), TextBox).Rows = System.Text.RegularExpressions.Regex.Matches(comment, "\r\n").Count + 1
                End If
                CType(SignT.FindControl("txt_comm_" & seq), TextBox).Text = comment
                If (status = 2 Or status = 100) Then
                    CType(SignT.FindControl("image_sign_" & seq), Image).ImageUrl = "~/image/ok1.jpg"
                ElseIf (status = 3 Or status = 5) Then
                    CType(SignT.FindControl("image_sign_" & seq), Image).ImageUrl = "~/image/rj1.jpg"
                ElseIf (status = 10) Then
                    CType(SignT.FindControl("image_sign_" & seq), Image).ImageUrl = "~/image/skip1.jpg"
                End If
            Loop
        End If
        dr.Close()
        connsap.Close()
        'End If
    End Sub
    Sub CreateHeadField()
        CLHead1()
        CLHead2()
        CLHead3()
        CLHead4()
        CLHead5()
        CLHead6()
        CLHead7()
        If (sfid = 2 Or sfid = 22) Then
            CLHead8()
            CLHead9()
        End If
        CLHead10()
        CLHead11()
        CLHead12()
        CLHeadTail()
    End Sub

    Sub CLHead1()
        Dim tCell As TableCell
        Dim tRow As TableRow
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "Jet內部聯絡單"
        tCell.BackColor = Drawing.Color.Gainsboro
        tCell.ColumnSpan = 10
        tCell.Font.Bold = True
        tCell.Font.Size = 24
        tRow.Cells.Add(tCell)
        HeadT.Rows.Add(tRow)
    End Sub
    Sub CLHead2()
        Dim tCell As TableCell
        Dim tRow As TableRow

        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "發文日期"
        tCell.BackColor = Drawing.Color.Gainsboro
        tCell.ColumnSpan = 2
        tCell.Font.Bold = True
        tCell.Font.Size = 12
        tRow.Cells.Add(tCell)
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        'tCell.Text = Format(Now(), "yyyy/MM/dd")
        'tCell.ColumnSpan = 2
        tCell.Font.Bold = True
        tCell.Font.Size = 12
        tRow.Cells.Add(tCell)

        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "發文單位"
        tCell.BackColor = Drawing.Color.Gainsboro
        'tCell.ColumnSpan = 3
        tCell.Font.Bold = True
        tCell.Font.Size = 12
        tRow.Cells.Add(tCell)

        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.ColumnSpan = 4
        RBLCDpt.ID = "rbl_cdpt"
        RBLCDpt.Items.Add("台北捷智TW")
        RBLCDpt.Items.Add("華東捷豐KS")
        RBLCDpt.Items.Add("華南捷智通SZ")
        RBLCDpt.Items(0).Value = 1
        RBLCDpt.Items(1).Value = 2
        RBLCDpt.Items(2).Value = 3
        RBLCDpt.RepeatDirection = RepeatDirection.Vertical
        RBLCDpt.AutoPostBack = True
        AddHandler RBLCDpt.SelectedIndexChanged, AddressOf RBLCDpt_SelectedIndexChanged
        tCell.Controls.Add(RBLCDpt)
        tRow.Cells.Add(tCell)

        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "聯絡單號"
        tCell.BackColor = Drawing.Color.Gainsboro
        'tCell.ColumnSpan = 4
        tCell.Font.Bold = True
        tCell.Font.Size = 12
        tRow.Cells.Add(tCell)

        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Font.Bold = True
        'tCell.ColumnSpan = 2
        tCell.Font.Size = 12
        tRow.Cells.Add(tCell)
        HeadT.Rows.Add(tRow)
    End Sub
    Sub CLHead3()
        Dim tCell As TableCell
        Dim tRow As TableRow
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "發文型式"
        tCell.BackColor = Drawing.Color.Gainsboro
        tCell.ColumnSpan = 4
        tCell.Font.Bold = True
        tCell.Font.Size = 12
        tRow.Cells.Add(tCell)
        '-
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "收文單位"
        tCell.BackColor = Drawing.Color.Gainsboro
        tCell.ColumnSpan = 6
        tCell.Font.Bold = True
        tCell.Font.Size = 12
        tRow.Cells.Add(tCell)
        HeadT.Rows.Add(tRow)
    End Sub
    Sub CLHead4()
        Dim tCell As TableCell
        Dim tRow As TableRow
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.ColumnSpan = 4
        RBLType.ID = "rbl_type"
        RBLType.Items.Add("申請")
        RBLType.Items.Add("通知")
        RBLType.Items.Add("聯絡")
        RBLType.Items.Add("報告")
        RBLType.Items.Add("其它")
        RBLType.Items(0).Value = 1
        RBLType.Items(1).Value = 2
        RBLType.Items(2).Value = 3
        RBLType.Items(3).Value = 4
        RBLType.Items(4).Value = 5
        RBLType.RepeatDirection = RepeatDirection.Vertical
        'AddHandler RBLType.SelectedIndexChanged, AddressOf RBLType_SelectedIndexChanged
        tCell.Controls.Add(RBLType)
        tRow.Cells.Add(tCell)
        '-
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.ColumnSpan = 6
        RBLDpt.ID = "rbl_dpt"
        RBLDpt.Items.Add("台北捷智TW")
        RBLDpt.Items.Add("華東捷豐KS")
        RBLDpt.Items.Add("華南捷智通SZ")
        RBLDpt.Items(0).Value = 1
        RBLDpt.Items(1).Value = 2
        RBLDpt.Items(2).Value = 3
        RBLDpt.RepeatDirection = RepeatDirection.Vertical
        'AddHandler RBLDpt.SelectedIndexChanged, AddressOf RBLDpt_SelectedIndexChanged
        tCell.Controls.Add(RBLDpt)
        tRow.Cells.Add(tCell)
        HeadT.Rows.Add(tRow)
    End Sub
    Sub CLHead5()
        Dim tCell As TableCell
        Dim tRow As TableRow
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "收文窗口"
        tCell.BackColor = Drawing.Color.Gainsboro
        tCell.ColumnSpan = 10
        tCell.Font.Bold = True
        tCell.Font.Size = 12
        tRow.Cells.Add(tCell)
        HeadT.Rows.Add(tRow)
    End Sub
    Sub CLHead6()
        Dim tCell As TableCell
        Dim tRow As TableRow
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.ColumnSpan = 10
        ChkDTDpt = New CheckBoxList
        ChkDTDpt.ID = "chk_dtdpt"
        ChkDTDpt.Items.Add("業務")
        ChkDTDpt.Items.Add("工程")
        ChkDTDpt.Items.Add("品保")
        ChkDTDpt.Items.Add("採購")
        ChkDTDpt.Items.Add("客服")
        ChkDTDpt.Items.Add("製造")
        ChkDTDpt.Items.Add("倉庫")
        ChkDTDpt.Items.Add("機構")
        ChkDTDpt.Items.Add("光學")
        ChkDTDpt.Items.Add("電控")
        ChkDTDpt.Items.Add("軟體")
        ChkDTDpt.Items.Add("電子")
        ChkDTDpt.Items.Add("其它")
        tCell.Controls.Add(ChkDTDpt)
        ChkDTDpt.RepeatDirection = RepeatDirection.Vertical
        tRow.Cells.Add(tCell)
        HeadT.Rows.Add(tRow)
    End Sub
    Sub CLHead7()
        Dim tCell As TableCell
        Dim tRow As TableRow
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "發文人"
        tCell.BackColor = Drawing.Color.Gainsboro
        tCell.ColumnSpan = 2
        tCell.Font.Bold = True
        tCell.Font.Size = 12
        tRow.Cells.Add(tCell)
        '-
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = Session("s_name")
        tCell.ColumnSpan = 8
        tCell.Font.Bold = True
        tCell.Font.Size = 12
        tRow.Cells.Add(tCell)
        HeadT.Rows.Add(tRow)
    End Sub
    Sub CLHead8()
        Dim tCell As TableCell
        Dim tRow As TableRow
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Left
        tCell.Text = "適用性質  (非關備品倉申請/領用帳務處理項目,無需勾選此欄位)"
        tCell.BackColor = Drawing.Color.Gainsboro
        tCell.ColumnSpan = 10
        tCell.Font.Bold = True
        tCell.Font.Size = 12
        tRow.Cells.Add(tCell)
        HeadT.Rows.Add(tRow)
    End Sub
    Sub CLHead9()
        Dim tCell As TableCell
        Dim tRow As TableRow
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.ColumnSpan = 10
        RBLSpareType = New RadioButtonList
        RBLSpareType.ID = "rbl_sparetype"
        RBLSpareType.Items.Add("申請備品(入庫)")
        RBLSpareType.Items.Add("領用備品(出庫)")
        RBLSpareType.Items.Add("申請備品並領用(入庫後即出庫)")
        RBLSpareType.Items.Add("改機換料(舊料寄回)")
        RBLSpareType.Items.Add("改機換料(舊料留備品)")
        RBLSpareType.Items(0).Value = 1
        RBLSpareType.Items(1).Value = 2
        RBLSpareType.Items(2).Value = 3
        RBLSpareType.Items(2).Value = 4
        RBLSpareType.Items(2).Value = 5
        tCell.Controls.Add(RBLSpareType)
        RBLSpareType.RepeatDirection = RepeatDirection.Vertical
        tRow.Cells.Add(tCell)
        HeadT.Rows.Add(tRow)

    End Sub
    Sub CLHead10()
        Dim tCell As TableCell
        Dim tRow As TableRow
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "發文主旨"
        tCell.BackColor = Drawing.Color.Gainsboro
        tCell.ColumnSpan = 2
        tCell.Font.Bold = True
        tCell.Font.Size = 12
        tRow.Cells.Add(tCell)
        '-
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Left
        tCell.ColumnSpan = 8
        TxtSubject = New TextBox
        TxtSubject.ID = "txt_subject"
        TxtSubject.Width = 900
        TxtSubject.Rows = 1
        tCell.Controls.Add(TxtSubject)
        tRow.Cells.Add(tCell)
        HeadT.Rows.Add(tRow)
    End Sub
    Sub CLHead11()
        Dim tCell As TableCell
        Dim tRow As TableRow
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Gainsboro
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "事由說明"
        tCell.ColumnSpan = 10
        tCell.Font.Bold = True
        tCell.Font.Size = 12
        tRow.Cells.Add(tCell)
        HeadT.Rows.Add(tRow)

    End Sub
    Sub CLHead12()
        Dim tCell As TableCell
        Dim tRow As TableRow
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Left
        tCell.ColumnSpan = 10
        TxtInfo = New TextBox
        TxtInfo.ID = "txt_info"
        TxtInfo.Width = 1200
        TxtInfo.TextMode = TextBoxMode.MultiLine
        TxtInfo.Rows = 30
        tCell.Controls.Add(TxtInfo)
        tRow.Cells.Add(tCell)
        HeadT.Rows.Add(tRow)
    End Sub
    Sub CLHeadTail()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim i As Integer
        tRow = New TableRow()
        tRow.BorderWidth = 1
        For i = 0 To 9
            tCell = New TableCell()
            tCell.BorderWidth = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            tCell.Font.Bold = True
            tCell.Font.Size = 12
            tCell.Width = 50
            tRow.Cells.Add(tCell)
        Next
        HeadT.Rows.Add(tRow)
    End Sub
    Protected Sub RBLCDpt_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim prefix As String
        prefix = ""
        If (RBLCDpt.SelectedValue = 1) Then 'TW
            prefix = "TW"
        ElseIf (RBLCDpt.SelectedValue = 2) Then 'KS
            prefix = "KS"
        ElseIf (RBLCDpt.SelectedValue = 3) Then 'SZ
            prefix = "SZ"
        End If
        HeadT.Rows(1).Cells(5).Text = GetCno(HeadT.Rows(1).Cells(1).Text, prefix)
        ViewState("docstatus") = docstatus
    End Sub
    Function GetCno(DateStr As Date, prefix As String)
        Dim cno As Long
        Dim cnostr As String
        Dim go As Boolean
        cnostr = ""
        go = True
        If (docstatus <> "A") Then
            SqlCmd = "Select issuedunit,cno From dbo.[@XSHWS] T0 where T0.docentry=" & docnum
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            If (RBLCDpt.SelectedValue = dr(0)) Then
                cnostr = dr(1)
                go = False
            End If
            dr.Close()
            connsap.Close()
        End If
        If (go) Then
            For i = 1 To 30
                cno = CLng(Format(DateStr, "yyMMdd")) * 100 + i
                cnostr = prefix & CStr(cno)
                SqlCmd = "Select count(*) From dbo.[@XSHWS] T0 where T0.cno='" & cnostr & "'"
                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                dr.Read()
                If (dr(0) = 0) Then
                    dr.Close()
                    connsap.Close()
                    Exit For
                End If
                dr.Close()
                connsap.Close()
            Next
        End If
        Return cnostr
    End Function
    Protected Sub BtnDel_Click(sender As Object, e As EventArgs)
        Dim beapproved As Boolean
        Dim signfinish As Boolean
        Dim str() As String
        beapproved = False
        '先check 是否此單已被覆核過
        SqlCmd = "Select status " &
                 "FROM dbo.[@XASCH] where docnum=" & docnum
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            dr.Read()
            If (dr(0) <> docstatus) Then
                beapproved = True
                CommUtil.ShowMsg(Me, "此單已被他人覆核過,不予處理,進入下一單")
            End If
        Else
            CommUtil.ShowMsg(Me, Session("s_id") & " " & ds.Tables(0).Rows(Session("startindex"))("docnum") & "簽核列表中在資料庫中找不到")
        End If
        dr.Close()
        connsap.Close()
        If (beapproved = False) Then
            For i = 1 To DDLAttaFile.Items.Count - 1
                str = Split(DDLAttaFile.Items(i).Value, "_")
                If (CLng(str(0)) = docnum) Then
                    IO.File.Delete(targetPath & DDLAttaFile.Items(i).Value)
                    IO.File.Delete(localsignoffformdir & DDLAttaFile.Items(i).Value)
                End If
            Next
            SqlCmd = "delete from [dbo].[@XASCH] where docnum=" & docnum
            CommUtil.SqlSapExecute("del", SqlCmd, connsap)
            connsap.Close()
            SqlCmd = "delete from [dbo].[@XRSCT] where docentry=" & docnum
            CommUtil.SqlSapExecute("del", SqlCmd, connsap)
            connsap.Close()
            SqlCmd = "delete from [dbo].[@XSPWT] where docentry=" & docnum
            CommUtil.SqlSapExecute("del", SqlCmd, connsap)
            connsap.Close()
            SqlCmd = "delete from [dbo].[@XSPHT] where docentry=" & docnum
            CommUtil.SqlSapExecute("del", SqlCmd, connsap)
            connsap.Close()
            SqlCmd = "delete from [dbo].[@XSMLS] where docentry=" & docnum
            CommUtil.SqlSapExecute("del", SqlCmd, connsap)
            connsap.Close()
            SqlCmd = "delete from [dbo].[@XRSCT] where docentry=" & docnum
            CommUtil.SqlSapExecute("del", SqlCmd, connsap)
            connsap.Close()
            SqlCmd = "delete from [dbo].[@XMSCT] where docentry=" & docnum
            CommUtil.SqlSapExecute("del", SqlCmd, connsap)
            connsap.Close()
            SqlCmd = "delete from [dbo].[@XGCT] where docentry=" & docnum
            CommUtil.SqlSapExecute("del", SqlCmd, connsap)
            connsap.Close()
            '將設定之返還數量歸0
            If (sfid = 101) Then
                If (TxtAttaDoc.Text <> "NA" And TxtAttaDoc.Text <> "") Then
                    SqlCmd = "update [dbo].[@XSMLS] set nowrtnqty=0 " &
                        " where docentry=" & CLng(TxtAttaDoc.Text) & " and head=0"
                    CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                    connsap.Close()
                End If
            End If
        End If
        Dim status As String
        If (actmode = "signoff" Or actmode = "recycle" Or actmode = "recycle_login" Or actmode = "signoff_login") Then
            ds.Tables(0).Rows(Session("startindex"))("signoffflag") = 1 '設定該員簽核表單list中 , 此單設為已簽
            Session("ds") = ds
            signfinish = FindNextSignDoc()
            status = ds.Tables(0).Rows(Session("startindex"))("status")
            If (signfinish = False) Then
                Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&actmode=recycle&status=" & status &
                                "&docnum=" & ds.Tables(0).Rows(Session("startindex"))("docnum") &
                                "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & ds.Tables(0).Rows(Session("startindex"))("sfid") & "&agnid=" & agnidG & "&signflowmode=" & signflowmode)
            Else
                If (actmode = "recycle_login" Or actmode = "signoff_login") Then
                    Response.Redirect("~/signoff/signoff.aspx?smid=sg&smode=1&act=signfinish&signflowmode=" & signflowmode)
                Else
                    Response.Redirect("~/usermgm/logout.aspx?act=signfinish")
                End If
            End If
        Else
            Response.Redirect("~/signoff/signoff.aspx?smid=sg&smode=1&act=del&signflowmode=" & signflowmode)
        End If
    End Sub
    Function XRSCTFieldCheck()
        Dim ok As Boolean
        ok = True
        If (CType(ContentT.FindControl("ddlrsuser_6_0"), DropDownList).SelectedIndex = 0) Then
            ok = False
        End If
        If (CType(ContentT.FindControl("txtv5usercode_6_1"), TextBox).Text = "") Then
            ok = False
        End If
        If (CType(ContentT.FindControl("ddlfromday_2_1"), DropDownList).SelectedIndex = 0) Then
            ok = False
        End If
        If (CType(ContentT.FindControl("ddltoday_2_1"), DropDownList).SelectedIndex = 0) Then
            ok = False
        End If
        If (CType(ContentT.FindControl("txtrsreason_4_1"), TextBox).Text = "") Then
            ok = False
        End If
        Return ok
    End Function
    Function XCMRTFieldCheck()
        Dim ok As Boolean
        ok = True
        If (CType(ContentT.FindControl("txt_reportdate"), TextBox).Text = "") Then
            ok = False
        End If
        If (CType(ContentT.FindControl("txt_machinetype"), TextBox).Text = "") Then
            ok = False
        End If
        If (CType(ContentT.FindControl("txt_cusname"), TextBox).Text = "") Then
            ok = False
        End If
        If (CType(ContentT.FindControl("txt_model"), TextBox).Text = "") Then
            ok = False
        End If
        If (CType(ContentT.FindControl("txt_machineserialOrwo"), TextBox).Text = "") Then
            ok = False
        End If
        If (CType(ContentT.FindControl("txt_faeperson"), TextBox).Text = "") Then
            ok = False
        End If
        If (CType(ContentT.FindControl("txt_problemdescrip"), TextBox).Text = "") Then
            ok = False
        End If
        If (CType(ContentT.FindControl("txt_processdescrip"), TextBox).Text = "") Then
            ok = False
        End If
        'If (CType(ContentT.FindControl("txt_verifydescrip"), TextBox).Text = "") Then
        'ok = False
        'End If
        Return ok
    End Function
    Function XFMRTFieldCheck()
        Dim ok As Boolean
        ok = True
        If (CType(ContentT.FindControl("txt_reportdate"), TextBox).Text = "") Then
            ok = False
        End If
        If (CType(ContentT.FindControl("txt_machinetype"), TextBox).Text = "") Then
            ok = False
        End If
        If (CType(ContentT.FindControl("txt_cusname"), TextBox).Text = "") Then
            ok = False
        End If
        If (CType(ContentT.FindControl("txt_model"), TextBox).Text = "") Then
            ok = False
        End If
        If (CType(ContentT.FindControl("txt_machineserialOrwo"), TextBox).Text = "") Then
            ok = False
        End If
        If (CType(ContentT.FindControl("txt_qcperson"), TextBox).Text = "") Then
            ok = False
        End If
        If (CType(ContentT.FindControl("txt_problemdescrip"), TextBox).Text = "") Then
            ok = False
        End If
        If (CType(ContentT.FindControl("txt_processdescrip"), TextBox).Text = "") Then
            ok = False
        End If
        'If (CType(ContentT.FindControl("txt_verifydescrip"), TextBox).Text = "") Then
        'ok = False
        'End If
        Return ok
    End Function
    Function XGCTFieldCheck()
        Dim ok As Boolean
        ok = True
        If (TxtReason.Text = "") Then
            ok = False
        End If
        If (TxtDept.Text = "") Then
            ok = False
        End If
        If (TxtPerson.Text = "") Then
            ok = False
        End If
        Return ok
    End Function
    Function XMSCTFieldCheck()
        Dim ok As Boolean
        ok = True
        If (CType(ContentT.FindControl("rbl_area"), RadioButtonList).SelectedIndex = -1) Then
            ok = False
        End If
        If (CType(ContentT.FindControl("txt_amount"), TextBox).Text = "") Then
            ok = False
        End If
        If (CType(ContentT.FindControl("rbl_plheight"), RadioButtonList).SelectedIndex = -1) Then
            ok = False
        End If
        If (CType(ContentT.FindControl("rbl_withfixture"), RadioButtonList).SelectedIndex = -1) Then
            ok = False
        End If
        If (CType(ContentT.FindControl("rbl_pcbdir"), RadioButtonList).SelectedIndex = -1) Then
            ok = False
        End If
        If (CType(ContentT.FindControl("rbl_oslang"), RadioButtonList).SelectedIndex = -1) Then
            ok = False
        End If
        If (CType(ContentT.FindControl("rbl_camerapixel"), RadioButtonList).SelectedIndex = -1) Then
            ok = False
        End If
        If (CType(ContentT.FindControl("rbl_rgb"), RadioButtonList).SelectedIndex = -1) Then
            ok = False
        End If
        If (CType(ContentT.FindControl("rbl_resolution"), RadioButtonList).SelectedIndex = -1) Then
            ok = False
        End If
        If (CType(ContentT.FindControl("rbl_rbgcontrol"), RadioButtonList).SelectedIndex = -1) Then
            ok = False
        End If
        If (CType(ContentT.FindControl("rbl_coaxialinstall"), RadioButtonList).SelectedIndex = -1) Then
            ok = False
        End If
        'If (CType(ContentT.FindControl("rbl_coaxialcolor"), RadioButtonList).SelectedIndex = -1) Then
        'ok = False
        'End If
        If (CType(ContentT.FindControl("rbl_belttype"), RadioButtonList).SelectedIndex = -1) Then
            ok = False
        End If
        If (CType(ContentT.FindControl("rbl_tblens"), RadioButtonList).SelectedIndex = -1) Then
            ok = False
        End If

        If (CType(ContentT.FindControl("txt_sales"), TextBox).Text = "") Then
            ok = False
        End If
        If (CType(ContentT.FindControl("ddl_model"), DropDownList).SelectedIndex = 0) Then
            ok = False
        End If
        If (CType(ContentT.FindControl("txt_customer"), TextBox).Text = "") Then
            ok = False
        End If
        If (CType(ContentT.FindControl("txt_shipmodel"), TextBox).Text = "") Then
            ok = False
        End If
        If (CType(ContentT.FindControl("txt_shipdate"), TextBox).Text = "") Then
            ok = False
        End If
        If (CType(ContentT.FindControl("txt_uutlength"), TextBox).Text = "") Then
            ok = False
        End If
        If (CType(ContentT.FindControl("txt_uutwidth"), TextBox).Text = "") Then
            ok = False
        End If
        'CType(ContentT.FindControl("txt_uutweight"), TextBox).Text = "") Then
        'CType(ContentT.FindControl("txt_uutthick"), TextBox).Text = "") Then
        If (CType(ContentT.FindControl("rbl_plheight"), RadioButtonList).SelectedIndex = 1) Then
            If (CType(ContentT.FindControl("txt_plotherheight"), TextBox).Text = "") Then
                ok = False
            End If
            If (CType(ContentT.FindControl("txt_plotherheighttol"), TextBox).Text = "") Then
                ok = False
            End If
        End If
        'If (CType(ContentT.FindControl("txt_fixturesize"), TextBox).Text = "") Then
        'ok = False
        'End If
        'CType(ContentT.FindControl("txt_pcbsizeX"), TextBox).Text = "") Then
        'CType(ContentT.FindControl("txt_pcbsizeY"), TextBox).Text = "") Then
        'CType(ContentT.FindControl("txt_cycletime"), TextBox).Text = "") Then
        'CType(ContentT.FindControl("txt_zmm"), TextBox).Text = "") Then
        If (CType(ContentT.FindControl("txt_topspace"), TextBox).Text = "") Then
            ok = False
            'CommUtil.ShowMsg(Me, "上方空間規格沒填")
        End If
        If (CType(ContentT.FindControl("txt_botspace"), TextBox).Text = "") Then
            ok = False
            'CommUtil.ShowMsg(Me, "下方空間規格沒填")
        End If
        If (CType(ContentT.FindControl("rbl_resolution"), RadioButtonList).SelectedIndex = 11) Then
            If (CType(ContentT.FindControl("txt_otherresolution"), TextBox).Text = "") Then
                ok = False
            End If
        End If
        If (CType(ContentT.FindControl("txt_memo"), TextBox).Text = "") Then
            ok = False
        End If
        Return ok
    End Function
    Protected Sub BtnSave_Click(sender As Object, e As EventArgs)
        Dim input_finish, mainattach As Boolean
        Dim subject, status, priceunit, attadoc As String
        Dim sapno As Long
        Dim price As Double
        Dim s_name, dept, area, spos As String
        Dim mcount As Integer
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim Now_time As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        priceunit = ""
        price = 0
        sapno = 0
        input_finish = True
        mainattach = True
        spos = ""
        If (TxtSubject.Text = "") Then
            input_finish = False
        End If
        If (sfid = 51 Or sfid = 50 Or sfid = 49 Or sfid = 23 Or sfid = 24) Then 'sfid process
            SqlCmd = "Select count(*) FROM [dbo].[@XSMLS] T0 WHERE head=0 and T0.[docentry] =" & docnum
            drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
            drL.Read()
            mcount = drL(0)
            If (mcount = 0) Then
                input_finish = False
            End If
            drL.Close()
            connL.Close()
        ElseIf (sfid = 16) Then 'sfid process
            If (Not XRSCTFieldCheck()) Then
                input_finish = False
            End If
        ElseIf (sfid = 12) Then 'sfid process
            If (Not XMSCTFieldCheck()) Then
                input_finish = False
            End If
        ElseIf (sfid = 1) Then 'sfid process
            If (Not XGCTFieldCheck()) Then
                input_finish = False
            End If
        ElseIf (sfid = 22) Then 'sfid process
            If (Not XCMRTFieldCheck()) Then
                input_finish = False
            End If
        ElseIf (sfid = 3) Then 'sfid process
            If (Not XFMRTFieldCheck()) Then
                input_finish = False
            End If
        End If
        If (sfid = 49 Or sfid = 50 Or sfid = 51) Then 'sfid process 這2個是 special
            If (TxtPrice.Text <> "" And TxtPrice.Text <> "NA" And TxtPrice.Text <> "0") Then
                price = CDbl(TxtPrice.Text)
            Else
                input_finish = False
            End If
            If (TxtReason.Text = "") Then
                input_finish = False
            End If
        ElseIf (sfid = 100) Then
            If (TxtReason.Text = "") Then
                input_finish = False
            End If
        ElseIf (sfid = 23 Or sfid = 24) Then
            If (TxtReason.Text.Length <= 60) Then
                input_finish = False
            End If
        End If
        If (sfid > 50 And sfid < 80) Then
            If (sfid > 70 And sfid < 80) Then
                If (TxtSapNO.Text <> "" And TxtSapNO.Text <> "NA" And TxtSapNO.Text <> "0") Then
                    sapno = CLng(TxtSapNO.Text)
                Else
                    input_finish = False
                End If
            End If
            If (TxtPrice.Text <> "" And TxtPrice.Text <> "NA" And TxtPrice.Text <> "0") Then
                price = CDbl(TxtPrice.Text)
            Else
                input_finish = False
            End If

            If (DDLDollorUnit.SelectedIndex <> 0) Then
                priceunit = DDLDollorUnit.SelectedValue
            Else
                input_finish = False
            End If
        End If
        attadoc = TxtAttaDoc.Text
        If (attadoc = "") Then
            input_finish = False
        ElseIf (attadoc <> "NA") Then
            If (Not IsNumeric(attadoc)) Then
                CommUtil.ShowMsg(Me, "被附屬單號要為整數數字")
                Exit Sub
            End If
            SqlCmd = "Select sum(nowrtnqty) FROM [dbo].[@XSMLS] T0 WHERE T0.[docentry] =" & CLng(attadoc)
            drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
            If (drL.HasRows) Then
                drL.Read()
                If (sfid = 101) Then
                    If (drL(0) = 0) Then
                        input_finish = False
                    End If
                End If
            End If
            drL.Close()
            connL.Close()
        End If
        Dim di As DirectoryInfo
        di = New DirectoryInfo(HttpContext.Current.Server.MapPath("~/") & "AttachFile\SignOffsFormFiles\" & sid_create & "\" & sfid & "\")
        Dim fi As FileInfo()
        If (CommSignOff.IsSelfForm(sfid) = 0) Then
            If (System.IO.Directory.Exists(HttpContext.Current.Server.MapPath("~/") & "AttachFile\SignOffsFormFiles\" & sid_create & "\" & sfid & "\")) Then
                fi = di.GetFiles(docnum & "*主檔*")
                If (fi.Length = 0) Then
                    mainattach = False
                End If
            Else
                mainattach = False
            End If
        End If
        If (input_finish And mainattach) Then
            If (docstatus = "E" Or docstatus = "A") Then
                status = "D"
            Else
                status = docstatus
            End If
        Else
            'status = docstatus
            If (docstatus = "A") Then
                status = "E"
            ElseIf (docstatus = "E") Then
                status = docstatus
            ElseIf (docstatus = "D") Then
                'CommUtil.ShowMsg(Me, "因本已完全,但你修改成不完全,還有欄位未填入或沒附檔, 故把狀態修改為編輯") '因之後導引至從頭開始,故此地方show不出
                status = "E"
                'Exit Sub
            End If
        End If
        subject = SubjectTextGen()
        If (docstatus <> "A") Then
            SqlCmd = "update [dbo].[@XASCH] set subject='" & subject & "',status='" & status & "', " &
                 "priceunit='" & priceunit & "',price=" & price & ",sapno=" & sapno & ",attadoc='" & attadoc & "' " &
                 " where docnum=" & docnum
            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
            connsap.Close()
            If (sfid = 16) Then 'sfid process
                UpdOrInsRecordSfid16("upd", docnum)
            ElseIf (sfid = 12) Then
                UpdOrInsRecordSfid12("upd", docnum)
            ElseIf (sfid = 49 Or sfid = 50 Or sfid = 51 Or sfid = 100) Then
                UpdOrInsRecordSfid49_50_51_100("upd", docnum)
            ElseIf (sfid = 1) Then
                UpdOrInsRecordSfid1("upd", docnum)
            ElseIf (sfid = 22) Then
                UpdOrInsRecordSfid22("upd", docnum)
            ElseIf (sfid = 3) Then
                UpdOrInsRecordSfid3("upd", docnum)
            ElseIf (sfid = 23 Or sfid = 24) Then
                UpdOrInsRecordSfid23_24("upd", docnum)
            End If
        Else
            s_name = HeadT.Rows(0).Cells(3).Text
            SqlCmd = "select position from dbo.[user] where id='" & sid & "'"
            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
            If (dr.HasRows) Then
                dr.Read()
                spos = dr(0)
            End If
            dr.Close()
            conn.Close()
            dept = HeadT.Rows(0).Cells(5).Text
            area = HeadT.Rows(0).Cells(7).Text
            SqlCmd = "insert into [dbo].[@XASCH] (docdate,sid,sname,sfid,dept,area,receivedate,spos,subject,status,priceunit,price,sapno,attadoc) " &
            "values(" & "'" & Now_time & "','" & sid & "','" & s_name & "'," & sfid & ",'" & dept & "','" & area & "','" & Now_time & "','" &
                    spos & "','" & subject & "','" & status & "','" & priceunit & "'," & price & "," & sapno & ",'" & attadoc & "')"
            CommUtil.SqlSapExecute("ins", SqlCmd, connsap)
            connsap.Close()
            SqlCmd = "Select max(docnum) from [dbo].[@XASCH] where sid='" & sid & "'"
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            docnum = dr(0)
            dr.Close()
            connsap.Close()
            'sfid=16時用
            If (sfid = 16) Then 'sfid process
                UpdOrInsRecordSfid16("ins", docnum) 'aaaaa
            ElseIf (sfid = 12) Then
                UpdOrInsRecordSfid12("ins", docnum)
            ElseIf (sfid = 49 Or sfid = 50 Or sfid = 51 Or sfid = 100) Then
                UpdOrInsRecordSfid49_50_51_100("ins", docnum)
            ElseIf (sfid = 1) Then
                UpdOrInsRecordSfid1("ins", docnum)
            ElseIf (sfid = 22) Then
                UpdOrInsRecordSfid22("ins", docnum)
            ElseIf (sfid = 3) Then
                UpdOrInsRecordSfid3("ins", docnum)
            ElseIf (sfid = 23 Or sfid = 24) Then
                UpdOrInsRecordSfid23_24("ins", docnum)
            End If
        End If
        If (input_finish And mainattach) Then
            If (sfid > 50 And sfid < 80 And (docstatus = "E" Or docstatus = "A")) Then
                Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&act=save&status=" & status & "&docnum=" & docnum &
                              "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & sfid & "&subject=" & TxtSubject.Text &
                                "&sapno=" & TxtSapNO.Text & "&price=" & TxtPrice.Text & "&unitindex=" & DDLDollorUnit.SelectedIndex &
                                "&signflowmode=" & signflowmode & "&attachsel=" & DDLAttaFile.SelectedValue & "&maindocnum=" & maindocnum & "&fromasp=" & fromasp)
            Else
                Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&act=save&status=" & status & "&docnum=" & docnum &
                              "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & sfid & "&subject=" & TxtSubject.Text &
                              "&signflowmode=" & signflowmode & "&attachsel=" & DDLAttaFile.SelectedValue & "&maindocnum=" & maindocnum & "&fromasp=" & fromasp)
            End If
        ElseIf (input_finish = True And mainattach = False) Then
            If (docstatus = "A") Then
                Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&act=create&status=" & status & "&docnum=" & docnum &
                              "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & sfid & "&subject=" & TxtSubject.Text &
                              "&signflowmode=" & signflowmode & "&attachsel=" & DDLAttaFile.SelectedValue & "&maindocnum=" & maindocnum & "&fromasp=" & fromasp)
            Else
                Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&act=notsaveattach&status=" & status & "&docnum=" & docnum &
                              "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & sfid & "&subject=" & TxtSubject.Text &
                              "&signflowmode=" & signflowmode & "&attachsel=" & DDLAttaFile.SelectedValue & "&maindocnum=" & maindocnum & "&fromasp=" & fromasp)
            End If
        Else
            If (docstatus = "A") Then
                Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&act=create&status=" & status & "&docnum=" & docnum &
                              "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & sfid & "&subject=" & TxtSubject.Text &
                              "&signflowmode=" & signflowmode & "&attachsel=" & DDLAttaFile.SelectedValue & "&maindocnum=" & maindocnum & "&fromasp=" & fromasp)
            Else
                Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&act=notsave&status=" & status & "&docnum=" & docnum &
                              "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & sfid & "&subject=" & TxtSubject.Text &
                              "&signflowmode=" & signflowmode & "&attachsel=" & DDLAttaFile.SelectedValue & "&maindocnum=" & maindocnum & "&fromasp=" & fromasp)
            End If
        End If
    End Sub
    Function SubjectTextGen()
        Dim idname, startdate, ddl_model, txt_customer, str(), subject As String
        Dim rbl_area, txt_amount As Integer
        idname = ""
        startdate = ""
        subject = ""
        subject = TxtSubject.Text
        If (sfid = 16) Then
            If (CType(ContentT.FindControl("ddlrsuser_6_0"), DropDownList).SelectedIndex <> 0) Then
                str = Split(CType(ContentT.FindControl("ddlrsuser_6_0"), DropDownList).SelectedValue, " ")
                idname = str(1)
            End If
            If (idname <> "") Then
                startdate = CType(ContentT.FindControl("ddlfromyear_2_1"), DropDownList).SelectedValue & "/" &
                    CType(ContentT.FindControl("ddlfrommonth_2_1"), DropDownList).SelectedValue & "/" &
                    CType(ContentT.FindControl("ddlfromday_2_1"), DropDownList).SelectedValue
                subject = idname & "申請 " & startdate & " 補刷卡"
                TxtSubject.Text = subject
            End If
        ElseIf (sfid = 12) Then
            rbl_area = CType(ContentT.FindControl("rbl_area"), RadioButtonList).SelectedIndex
            ddl_model = Split(CType(ContentT.FindControl("ddl_model"), DropDownList).SelectedValue, "-")(0)
            txt_customer = CType(ContentT.FindControl("txt_customer"), TextBox).Text
            If (IsNumeric(CType(ContentT.FindControl("txt_amount"), TextBox).Text)) Then
                txt_amount = CType(ContentT.FindControl("txt_amount"), TextBox).Text
            Else
                txt_amount = 0
            End If
            If (rbl_area <> -1 And ddl_model <> "" And txt_customer <> "" And txt_amount <> 0) Then
                If (rbl_area = 0) Then
                    subject = "台北發包:" & ddl_model & "機台 數量:" & txt_amount & "台 客戶:" & txt_customer
                ElseIf (rbl_area = 1) Then
                    subject = "捷智通發包:" & ddl_model & "機台 數量:" & txt_amount & "台 客戶:" & txt_customer
                ElseIf (rbl_area = 2) Then
                    subject = "捷豐發包:" & ddl_model & "機台 數量:" & txt_amount & "台 客戶:" & txt_customer
                End If
                TxtSubject.Text = subject
            End If
        End If
        Return subject
    End Function
    Sub UpdOrInsRecordSfid22(mode As String, docnum As Long)
        Dim reportdate, machinetype, cusname, cusfactory, model, machineserial, installdate As String
        Dim firstinstall, inwarranty As Integer
        Dim problemtype, typedescrip, verandspec, faeperson, problemdescrip, processdescrip, verifydescrip, problemnote As String
        firstinstall = 0
        inwarranty = 0
        reportdate = CType(ContentT.FindControl("txt_reportdate"), TextBox).Text
        machinetype = CType(ContentT.FindControl("txt_machinetype"), TextBox).Text
        cusname = CType(ContentT.FindControl("txt_cusname"), TextBox).Text
        cusfactory = CType(ContentT.FindControl("txt_cusfactoryOrmo"), TextBox).Text
        model = CType(ContentT.FindControl("txt_model"), TextBox).Text
        machineserial = CType(ContentT.FindControl("txt_machineserialOrwo"), TextBox).Text
        installdate = CType(ContentT.FindControl("txt_installdateOrshipdate"), TextBox).Text
        problemtype = CType(ContentT.FindControl("txt_problemtype"), TextBox).Text
        typedescrip = CType(ContentT.FindControl("txt_typedescrip"), TextBox).Text
        verandspec = CType(ContentT.FindControl("txt_verandspec"), TextBox).Text
        faeperson = CType(ContentT.FindControl("txt_faeperson"), TextBox).Text
        problemdescrip = CType(ContentT.FindControl("txt_problemdescrip"), TextBox).Text
        processdescrip = CType(ContentT.FindControl("txt_processdescrip"), TextBox).Text
        verifydescrip = CType(ContentT.FindControl("txt_verifydescrip"), TextBox).Text
        problemnote = CType(ContentT.FindControl("txt_problemnote"), TextBox).Text
        If (CType(ContentT.FindControl("chk_firstinstallOrnoassign"), CheckBox).Checked) Then
            firstinstall = 1
        End If
        If (CType(ContentT.FindControl("chk_inwarranty"), CheckBox).Checked) Then
            inwarranty = 1
        End If
        If (mode = "upd") Then
            SqlCmd = "update [dbo].[@XCMRT] set firstinstallOrnoassign=" & firstinstall & ",inwarranty=" & inwarranty & "," &
                    "reportdate='" & reportdate & "',machinetype='" & machinetype & "',cusname='" & cusname & "'," &
                    "cusfactoryOrmo='" & cusfactory & "',model='" & model & "',machineserialOrwo='" & machineserial & "'," &
                    "installdateOrshipdate='" & installdate & "',problemtype='" & problemtype & "',typedescrip='" & typedescrip & "'," &
                    "verandspec='" & verandspec & "',faeperson='" & faeperson & "',problemdescrip='" & problemdescrip & "'," &
                    "processdescrip='" & processdescrip & "',verifydescrip='" & verifydescrip & "',problemnote='" & problemnote & "' " &
                    "where docentry=" & docnum
            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
            connsap.Close()
        Else
            SqlCmd = "insert into [dbo].[@XCMRT] (firstinstallOrnoassign,inwarranty,reportdate,machinetype,cusname,cusfactoryOrmo,model," &
                                                "machineserialOrwo,installdateOrshipdate,problemtype,typedescrip,verandspec,faeperson," &
                                                "problemdescrip,processdescrip,verifydescrip,problemnote,docentry) " &
                     "values(" & firstinstall & "," & inwarranty & ",'" & reportdate & "','" & machinetype & "','" & cusname & "','" &
                            cusfactory & "','" & model & "','" & machineserial & "','" & installdate & "','" & problemtype & "','" &
                            typedescrip & "','" & verandspec & "','" & faeperson & "','" & problemdescrip & "','" & processdescrip & "','" &
                            verifydescrip & "','" & problemnote & "'," & docnum & ")"
            CommUtil.SqlSapExecute("ins", SqlCmd, connsap)
            connsap.Close()
        End If
    End Sub
    Sub UpdOrInsRecordSfid3(mode As String, docnum As Long)
        Dim reportdate, machinetype, cusname, mo, model, wo, shipdate As String
        Dim noassign As Integer
        Dim problemtype, typedescrip, verandspec, qcperson, problemdescrip, processdescrip, verifydescrip, problemnote As String
        noassign = 0
        reportdate = CType(ContentT.FindControl("txt_reportdate"), TextBox).Text
        machinetype = CType(ContentT.FindControl("txt_machinetype"), TextBox).Text
        cusname = CType(ContentT.FindControl("txt_cusname"), TextBox).Text
        mo = CType(ContentT.FindControl("txt_cusfactoryOrmo"), TextBox).Text
        model = CType(ContentT.FindControl("txt_model"), TextBox).Text
        wo = CType(ContentT.FindControl("txt_machineserialOrwo"), TextBox).Text
        shipdate = CType(ContentT.FindControl("txt_installdateOrshipdate"), TextBox).Text
        problemtype = CType(ContentT.FindControl("txt_problemtype"), TextBox).Text
        typedescrip = CType(ContentT.FindControl("txt_typedescrip"), TextBox).Text
        verandspec = CType(ContentT.FindControl("txt_verandspec"), TextBox).Text
        qcperson = CType(ContentT.FindControl("txt_qcperson"), TextBox).Text
        problemdescrip = CType(ContentT.FindControl("txt_problemdescrip"), TextBox).Text
        processdescrip = CType(ContentT.FindControl("txt_processdescrip"), TextBox).Text
        verifydescrip = CType(ContentT.FindControl("txt_verifydescrip"), TextBox).Text
        problemnote = CType(ContentT.FindControl("txt_problemnote"), TextBox).Text
        If (CType(ContentT.FindControl("chk_firstinstallOrnoassign"), CheckBox).Checked) Then
            noassign = 1
        End If
        If (mode = "upd") Then
            SqlCmd = "update [dbo].[@XCMRT] set firstinstallOrnoassign=" & noassign & "," &
                    "reportdate='" & reportdate & "',machinetype='" & machinetype & "',cusname='" & cusname & "'," &
                    "cusfactoryOrmo='" & mo & "',model='" & model & "',machineserialOrwo='" & wo & "'," &
                    "installdateOrshipdate='" & shipdate & "',problemtype='" & problemtype & "',typedescrip='" & typedescrip & "'," &
                    "verandspec='" & verandspec & "',qcperson='" & qcperson & "',problemdescrip='" & problemdescrip & "'," &
                    "processdescrip='" & processdescrip & "',verifydescrip='" & verifydescrip & "',problemnote='" & problemnote & "' " &
                    "where docentry=" & docnum
            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
            connsap.Close()
        Else
            SqlCmd = "insert into [dbo].[@XCMRT] (firstinstallOrnoassign,reportdate,machinetype,cusname,cusfactoryOrmo,model," &
                                                "machineserialOrwo,installdateOrshipdate,problemtype,typedescrip,verandspec,qcperson," &
                                                "problemdescrip,processdescrip,verifydescrip,problemnote,docentry) " &
                     "values(" & noassign & ",'" & reportdate & "','" & machinetype & "','" & cusname & "','" &
                            mo & "','" & model & "','" & wo & "','" & shipdate & "','" & problemtype & "','" &
                            typedescrip & "','" & verandspec & "','" & qcperson & "','" & problemdescrip & "','" & processdescrip & "','" &
                            verifydescrip & "','" & problemnote & "'," & docnum & ")"
            CommUtil.SqlSapExecute("ins", SqlCmd, connsap)
            connsap.Close()
        End If
    End Sub
    Sub UpdOrInsRecordSfid1(mode As String, docnum As Long)
        Dim txtctdept, txtctperson, txtctdescrip As String
        txtctdept = TxtDept.Text
        txtctperson = TxtPerson.Text
        txtctdescrip = TxtReason.Text
        If (mode = "upd") Then
            SqlCmd = "update [dbo].[@XGCT] set ctdept='" & txtctdept & "',ctperson='" & txtctperson & "'," &
                    "ctdescrip='" & txtctdescrip & "' " &
                    "where docentry=" & docnum
            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
            connsap.Close()
        Else
            SqlCmd = "insert into [dbo].[@XGCT] (ctdept,ctperson,ctdescrip,docentry) " &
                     "values(" & "'" & txtctdept & "','" & txtctperson & "','" & txtctdescrip & "'," & docnum & ")"
            CommUtil.SqlSapExecute("ins", SqlCmd, connsap)
            connsap.Close()
        End If
    End Sub
    Sub UpdOrInsRecordSfid49_50_51_100(mode As String, docnum As Long)
        If (mode = "upd") Then
            SqlCmd = "update [dbo].[@XSMLS] set descrip='" & TxtReason.Text & "'" &
                    " where docentry=" & docnum & " and head=1"
            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
            connsap.Close()
        Else
            Dim headexist As Boolean
            Dim descrip As String
            headexist = False
            descrip = ""
            SqlCmd = "select count(*) from [dbo].[@XSMLS] where docentry=" & docnum
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                dr.Read()
                If (dr(0) > 0) Then '如果>0就表示head (head=1) 已建立過
                    headexist = True
                End If
            End If
            dr.Close()
            connsap.Close()
            If (sfid = 50) Then
                descrip = "1.服務對象資訊(細分到廠區別):" & vbCrLf & vbCrLf & "2.設備資訊(型號/序號/裝機日期/驗收日期/保固與否/保固期限) :" & vbCrLf & vbCrLf &
                      "3.為什麼要申請此物料 ? :" & vbCrLf
            End If
            If (headexist = False) Then
                SqlCmd = "insert into [dbo].[@XSMLS] (docentry,head,descrip) " &
                        "values(" & docnum & ", 1,'" & descrip & "')"
                CommUtil.SqlSapExecute("ins", SqlCmd, connsap)
                connsap.Close()
            End If
        End If
    End Sub
    Sub UpdOrInsRecordSfid23_24(mode As String, docnum As Long)
        Dim headexist As Boolean
        Dim descrip As String
        Dim mtype As Integer
        'Dim drL As SqlDataReader
        'Dim connL As New SqlConnection
        headexist = False
        descrip = ""
        If (mode = "upd") Then
            SqlCmd = "update [dbo].[@XSMLS] set descrip='" & TxtReason.Text & "'" &
                    " where docentry=" & docnum & " and head=1"
            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
            connsap.Close()
        Else
            If (sfid = 23) Then
                descrip = "1.何單位(廠商)借用:" & vbCrLf & vbCrLf & "2.何人(代領)領取 :" & vbCrLf & vbCrLf &
                      "3.為什麼要借用此物料 ? :" & vbCrLf
                mtype = 1
            ElseIf (sfid = 24) Then
                descrip = "1.從何單位(客戶)借入:" & vbCrLf & vbCrLf & "2.何人(代給)給予 :" & vbCrLf & vbCrLf &
                      "3.為什麼要借入此物料 ? :" & vbCrLf
                mtype = 2
            End If
            SqlCmd = "select count(*) from [dbo].[@XSMLS] where docentry=" & docnum
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                dr.Read()
                If (dr(0) > 0) Then '如果>0就表示head (head=1) 已建立過
                    headexist = True
                End If
            End If
            dr.Close()
            connsap.Close()
            If (headexist = False) Then
                SqlCmd = "insert into [dbo].[@XSMLS] (docentry,head,descrip,mtype) " &
                            "values(" & docnum & ", 1,'" & descrip & "'," & mtype & ")"
                CommUtil.SqlSapExecute("ins", SqlCmd, connsap)
                connsap.Close()
            End If
        End If
    End Sub
    Sub UpdOrInsRecordSfid12(mode As String, docnum As Long)
        Dim rbl_area, txt_amount, rbl_plheight, rbl_withfixture, rbl_pcbdir, rbl_oslang, rbl_camerapixel, rbl_rgb As Integer
        Dim rbl_resolution, rbl_rbgcontrol, rbl_coaxialinstall, rbl_coaxialcolor, rbl_belttype, chk_flux, chk_anti As Integer
        Dim txt_sales, ddl_model, txt_customer, txt_shipmodel, txt_shipdate, txt_uutlength, txt_uutwidth, txt_uutweight As String
        Dim txt_uutthick, txt_plotherheight, txt_plotherheighttol, txt_fixturesize, txt_pcbsizeX, txt_pcbsizeY, txt_cycletime As String
        Dim txt_zmm, txt_topspace, txt_botspace, txt_otherresolution, txt_memo As String
        Dim str() As String
        Dim chk_sidecamera, rbl_tblens, chk_upz, chk_downz As Integer
        Dim txt_dzmm As String
        rbl_area = CType(ContentT.FindControl("rbl_area"), RadioButtonList).SelectedIndex
        rbl_plheight = CType(ContentT.FindControl("rbl_plheight"), RadioButtonList).SelectedIndex
        rbl_withfixture = CType(ContentT.FindControl("rbl_withfixture"), RadioButtonList).SelectedIndex
        rbl_pcbdir = CType(ContentT.FindControl("rbl_pcbdir"), RadioButtonList).SelectedIndex
        rbl_oslang = CType(ContentT.FindControl("rbl_oslang"), RadioButtonList).SelectedIndex
        rbl_camerapixel = CType(ContentT.FindControl("rbl_camerapixel"), RadioButtonList).SelectedIndex
        rbl_rgb = CType(ContentT.FindControl("rbl_rgb"), RadioButtonList).SelectedIndex
        rbl_resolution = CType(ContentT.FindControl("rbl_resolution"), RadioButtonList).SelectedIndex
        rbl_rbgcontrol = CType(ContentT.FindControl("rbl_rbgcontrol"), RadioButtonList).SelectedIndex
        rbl_coaxialinstall = CType(ContentT.FindControl("rbl_coaxialinstall"), RadioButtonList).SelectedIndex
        If (CType(ContentT.FindControl("rbl_coaxialinstall"), RadioButtonList).SelectedIndex = 0) Then
            rbl_coaxialcolor = CType(ContentT.FindControl("rbl_coaxialcolor"), RadioButtonList).SelectedIndex
        Else
            rbl_coaxialcolor = -1
        End If
        rbl_belttype = CType(ContentT.FindControl("rbl_belttype"), RadioButtonList).SelectedIndex
        rbl_tblens = CType(ContentT.FindControl("rbl_tblens"), RadioButtonList).SelectedIndex

        If (IsNumeric(CType(ContentT.FindControl("txt_amount"), TextBox).Text)) Then
            txt_amount = CType(ContentT.FindControl("txt_amount"), TextBox).Text
        Else
            txt_amount = 0
        End If
        chk_flux = CType(ContentT.FindControl("chk_flux"), CheckBox).Checked
        chk_anti = CType(ContentT.FindControl("chk_anti"), CheckBox).Checked
        chk_sidecamera = CType(ContentT.FindControl("chk_sidecamera"), CheckBox).Checked
        chk_upz = CType(ContentT.FindControl("chk_upz"), CheckBox).Checked
        chk_downz = CType(ContentT.FindControl("chk_downz"), CheckBox).Checked

        txt_sales = CType(ContentT.FindControl("txt_sales"), TextBox).Text
        str = Split(CType(ContentT.FindControl("ddl_model"), DropDownList).SelectedValue, "-")
        ddl_model = str(0)
        txt_customer = CType(ContentT.FindControl("txt_customer"), TextBox).Text
        txt_shipmodel = CType(ContentT.FindControl("txt_shipmodel"), TextBox).Text
        txt_shipdate = CType(ContentT.FindControl("txt_shipdate"), TextBox).Text
        txt_uutlength = CType(ContentT.FindControl("txt_uutlength"), TextBox).Text
        txt_uutwidth = CType(ContentT.FindControl("txt_uutwidth"), TextBox).Text
        txt_uutweight = CType(ContentT.FindControl("txt_uutweight"), TextBox).Text
        txt_uutthick = CType(ContentT.FindControl("txt_uutthick"), TextBox).Text
        If (CType(ContentT.FindControl("rbl_plheight"), RadioButtonList).SelectedIndex = 1) Then
            txt_plotherheight = CType(ContentT.FindControl("txt_plotherheight"), TextBox).Text
            txt_plotherheighttol = CType(ContentT.FindControl("txt_plotherheighttol"), TextBox).Text
        Else
            txt_plotherheight = ""
            txt_plotherheighttol = ""
        End If
        txt_fixturesize = CType(ContentT.FindControl("txt_fixturesize"), TextBox).Text
        txt_pcbsizeX = CType(ContentT.FindControl("txt_pcbsizeX"), TextBox).Text
        txt_pcbsizeY = CType(ContentT.FindControl("txt_pcbsizeY"), TextBox).Text
        txt_cycletime = CType(ContentT.FindControl("txt_cycletime"), TextBox).Text
        txt_zmm = CType(ContentT.FindControl("txt_zmm"), TextBox).Text
        txt_dzmm = CType(ContentT.FindControl("txt_dzmm"), TextBox).Text
        txt_topspace = CType(ContentT.FindControl("txt_topspace"), TextBox).Text
        txt_botspace = CType(ContentT.FindControl("txt_botspace"), TextBox).Text
        If (CType(ContentT.FindControl("rbl_resolution"), RadioButtonList).SelectedIndex = 11) Then
            txt_otherresolution = CType(ContentT.FindControl("txt_otherresolution"), TextBox).Text
        Else
            txt_otherresolution = ""
        End If
        txt_memo = CType(ContentT.FindControl("txt_memo"), TextBox).Text
        If (mode = "upd") Then
            SqlCmd = "update [dbo].[@XMSCT] set rbl_area=" & rbl_area & ",txt_amount=" & txt_amount & "," &
                    "rbl_plheight=" & rbl_plheight & ",rbl_withfixture=" & rbl_withfixture & ",rbl_pcbdir=" & rbl_pcbdir & "," &
                    "rbl_oslang=" & rbl_oslang & ",rbl_camerapixel=" & rbl_camerapixel & ",rbl_rgb=" & rbl_rgb & "," &
                    "rbl_resolution=" & rbl_resolution & ",rbl_rbgcontrol=" & rbl_rbgcontrol & ",rbl_coaxialinstall=" & rbl_coaxialinstall & "," &
                    "rbl_coaxialcolor=" & rbl_coaxialcolor & ",rbl_belttype=" & rbl_belttype & ",chk_flux=" & chk_flux & "," &
                    "chk_anti=" & chk_anti & ",txt_sales='" & txt_sales & "',ddl_model='" & ddl_model & "',txt_customer='" & txt_customer & "'," &
                    "txt_shipmodel='" & txt_shipmodel & "',txt_shipdate='" & txt_shipdate & "',txt_uutlength='" & txt_uutlength & "'," &
                    "txt_uutwidth='" & txt_uutwidth & "',txt_uutweight='" & txt_uutweight & "',txt_uutthick='" & txt_uutthick & "'," &
                    "txt_plotherheight='" & txt_plotherheight & "',txt_plotherheighttol='" & txt_plotherheighttol & "'," &
                    "txt_fixturesize='" & txt_fixturesize & "',txt_pcbsizeX='" & txt_pcbsizeX & "',txt_pcbsizeY='" & txt_pcbsizeY & "'," &
                    "txt_cycletime='" & txt_cycletime & "',txt_zmm='" & txt_zmm & "',txt_topspace='" & txt_topspace & "'," &
                    "txt_botspace='" & txt_botspace & "',txt_otherresolution='" & txt_otherresolution & "',txt_memo='" & txt_memo & "'," &
                    "txt_dzmm='" & txt_dzmm & "',rbl_tblens=" & rbl_tblens & ",chk_sidecamera=" & chk_sidecamera & "," &
                    "chk_upz=" & chk_upz & ",chk_downz=" & chk_downz &
                    " where docentry=" & docnum
            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
            connsap.Close()
        Else
            SqlCmd = "insert into [dbo].[@XMSCT] (docentry,rbl_area,txt_amount,rbl_plheight,rbl_withfixture,rbl_pcbdir,rbl_oslang," &
         "rbl_camerapixel,rbl_rgb,rbl_resolution,rbl_rbgcontrol,rbl_coaxialinstall,rbl_coaxialcolor,rbl_belttype," &
         "chk_flux,chk_anti,txt_sales,ddl_model,txt_customer,txt_shipmodel,txt_shipdate,txt_uutlength,txt_uutwidth," &
         "txt_uutweight,txt_uutthick,txt_plotherheight,txt_plotherheighttol,txt_fixturesize, txt_pcbsizeX, txt_pcbsizeY," &
         "txt_cycletime,txt_zmm, txt_topspace, txt_botspace, txt_otherresolution, txt_memo,txt_dzmm,rbl_tblens,chk_sidecamera,chk_upz,chk_downz) " &
         "values(" & docnum & "," & rbl_area & "," & txt_amount & "," & rbl_plheight & "," & rbl_withfixture & "," &
         rbl_pcbdir & "," & rbl_oslang & "," & rbl_camerapixel & "," & rbl_rgb & "," & rbl_resolution & "," &
         rbl_rbgcontrol & "," & rbl_coaxialinstall & "," & rbl_coaxialcolor & "," & rbl_belttype & "," &
         chk_flux & "," & chk_anti & ",'" & txt_sales & "','" & ddl_model & "','" & txt_customer & "','" &
         txt_shipmodel & "','" & txt_shipdate & "','" & txt_uutlength & "','" & txt_uutwidth & "','" &
         txt_uutweight & "','" & txt_uutthick & "','" & txt_plotherheight & "','" & txt_plotherheighttol & "','" &
         txt_fixturesize & "','" & txt_pcbsizeX & "','" & txt_pcbsizeY & "','" & txt_cycletime & "','" &
         txt_zmm & "','" & txt_topspace & "','" & txt_botspace & "','" & txt_otherresolution & "','" &
         txt_memo & "','" & txt_dzmm & "'," & rbl_tblens & "," & chk_sidecamera & "," & chk_upz & "," & chk_downz & ")"
            CommUtil.SqlSapExecute("ins", SqlCmd, connsap)
            connsap.Close()
        End If

    End Sub
    Sub UpdOrInsRecordSfid16(mode As String, docnum As Long)
        Dim id, idname, createid, createname, builtdate, startdate, bhour, bmin, enddate, ehour, emin, reason, v5id, str() As String
        id = ""
        idname = ""
        startdate = ""
        enddate = ""
        reason = ""
        v5id = ""
        If (CType(ContentT.FindControl("ddlrsuser_6_0"), DropDownList).SelectedIndex <> 0) Then
            str = Split(CType(ContentT.FindControl("ddlrsuser_6_0"), DropDownList).SelectedValue, " ")
            id = str(0)
            idname = str(1)
        End If
        builtdate = Format(Now(), "yyyy/MM/dd HH:mm:ss")
        If (CType(ContentT.FindControl("ddlfromday_2_1"), DropDownList).SelectedIndex <> 0) Then
            startdate = CType(ContentT.FindControl("ddlfromyear_2_1"), DropDownList).SelectedValue & "/" &
                        CType(ContentT.FindControl("ddlfrommonth_2_1"), DropDownList).SelectedValue & "/" &
                        CType(ContentT.FindControl("ddlfromday_2_1"), DropDownList).SelectedValue
        End If
        bhour = CType(ContentT.FindControl("ddlfromhour_3_1"), DropDownList).SelectedValue
        bmin = CType(ContentT.FindControl("ddlfrommin_3_1"), DropDownList).SelectedValue
        If (CType(ContentT.FindControl("ddltoday_2_1"), DropDownList).SelectedIndex <> 0) Then
            enddate = CType(ContentT.FindControl("ddltoyear_2_1"), DropDownList).SelectedValue & "/" &
                        CType(ContentT.FindControl("ddltomonth_2_1"), DropDownList).SelectedValue & "/" &
                        CType(ContentT.FindControl("ddltoday_2_1"), DropDownList).SelectedValue
        End If
        ehour = CType(ContentT.FindControl("ddltohour_3_1"), DropDownList).SelectedValue
        emin = CType(ContentT.FindControl("ddltomin_3_1"), DropDownList).SelectedValue
        createname = Session("s_name")
        If (CType(ContentT.FindControl("txtrsreason_4_1"), TextBox).Text <> "") Then
            reason = CType(ContentT.FindControl("txtrsreason_4_1"), TextBox).Text
        End If
        If (CType(ContentT.FindControl("txtv5usercode_6_1"), TextBox).Text <> "") Then
            v5id = CType(ContentT.FindControl("txtv5usercode_6_1"), TextBox).Text
        End If
        createid = Session("s_id")
        If (mode = "upd") Then
            If (startdate <> "" And enddate <> "") Then
                SqlCmd = "update [dbo].[@XRSCT] set id='" & id & "',idname='" & idname & "'," &
             "albdate='" & startdate & "',albhour=" & bhour & ",albmin=" & bmin & ",aledate='" & enddate & "'," &
             "alehour=" & ehour & ",alemin=" & emin & ",rsreason='" & reason & "',v5id='" & v5id & "'" &
             " where docentry=" & docnum
            ElseIf (startdate = "" And enddate <> "") Then
                SqlCmd = "update [dbo].[@XRSCT] set id='" & id & "',idname='" & idname & "'," &
             "albhour=" & bhour & ",albmin=" & bmin & ",aledate='" & enddate & "'," &
             "alehour=" & ehour & ",alemin=" & emin & ",rsreason='" & reason & "',v5id='" & v5id & "'" &
             " where docentry=" & docnum
            ElseIf (startdate <> "" And enddate = "") Then
                SqlCmd = "update [dbo].[@XRSCT] set id='" & id & "',idname='" & idname & "'," &
             "albdate='" & startdate & "',albhour=" & bhour & ",albmin=" & bmin & "," &
             "alehour=" & ehour & ",alemin=" & emin & ",rsreason='" & reason & "',v5id='" & v5id & "'" &
             " where docentry=" & docnum
            ElseIf (startdate = "" And enddate = "") Then
                SqlCmd = "update [dbo].[@XRSCT] set id='" & id & "',idname='" & idname & "'," &
             "albhour=" & bhour & ",albmin=" & bmin & "," &
             "alehour=" & ehour & ",alemin=" & emin & ",rsreason='" & reason & "',v5id='" & v5id & "'" &
             " where docentry=" & docnum
            End If
            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
            connsap.Close()
        Else
            If (startdate <> "" And enddate <> "") Then
                SqlCmd = "insert into [dbo].[@XRSCT] (id,idname,cdate,albdate,albhour,albmin,aledate,alehour,alemin,createid,createname,rsreason,v5id,docentry) " &
                        "values(" & "'" & id & "','" & idname & "','" & builtdate & "','" & startdate & "'," & bhour & "," & bmin & ",'" & enddate & "'," &
                        ehour & "," & emin & ",'" & createid & "','" & createname & "','" & reason & "','" & v5id & "'," & docnum & ")"
            ElseIf (startdate = "" And enddate <> "") Then
                SqlCmd = "insert into [dbo].[@XRSCT] (id,idname,cdate,albhour,albmin,aledate,alehour,alemin,createid,createname,rsreason,v5id,docentry) " &
                        "values(" & "'" & id & "','" & idname & "','" & builtdate & "'," & bhour & "," & bmin & ",'" & enddate & "'," &
                        ehour & "," & emin & ",'" & createid & "','" & createname & "','" & reason & "','" & v5id & "'," & docnum & ")"
            ElseIf (startdate <> "" And enddate = "") Then
                SqlCmd = "insert into [dbo].[@XRSCT] (id,idname,cdate,albdate,albhour,albmin,alehour,alemin,createid,createname,rsreason,v5id,docentry) " &
                        "values(" & "'" & id & "','" & idname & "','" & builtdate & "','" & startdate & "'," & bhour & "," & bmin & "," &
                        ehour & "," & emin & ",'" & createid & "','" & createname & "','" & reason & "','" & v5id & "'," & docnum & ")"
            ElseIf (startdate = "" And enddate = "") Then
                SqlCmd = "insert into [dbo].[@XRSCT] (id,idname,cdate,albhour,albmin,alehour,alemin,createid,createname,rsreason,v5id,docentry) " &
                        "values(" & "'" & id & "','" & idname & "','" & builtdate & "'," & bhour & "," & bmin & "," &
                        ehour & "," & emin & ",'" & createid & "','" & createname & "','" & reason & "','" & v5id & "'," & docnum & ")"
            End If
            CommUtil.SqlSapExecute("ins", SqlCmd, connsap)
            connsap.Close()
        End If
    End Sub
    Protected Sub BtnSend_Click(sender As Object, e As EventArgs)
        '修改狀態為O
        '產生簽核流程表
        '寫入送審歷史記錄
        Dim sftype As Integer
        Dim comment, status, historystatus As String
        Dim stoponlevel As Boolean
        Dim sapno As Long
        Dim beapproved, signfinish As Boolean
        'Dim maxseq As Integer
        Dim ownid As String
        ownid = ""
        beapproved = False
        historystatus = ""
        '先check 是否此單已被覆核過
        SqlCmd = "Select status " &
                 "FROM dbo.[@XASCH] where docnum=" & docnum
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            dr.Read()
            If (dr(0) <> docstatus) Then
                beapproved = True
                CommUtil.ShowMsg(Me, "此單已被他人覆核過,不予處理,進入下一單")
            End If
        Else
            CommUtil.ShowMsg(Me, Session("s_id") & " " & ds.Tables(0).Rows(Session("startindex"))("docnum") & "簽核列表中在資料庫中找不到")
        End If
        dr.Close()
        connsap.Close()
        '''''''''''''''''end
        If (beapproved = False) Then
            comment = TxtComm.Text
            SqlCmd = "Select subject,sapno,price " &
                "from [dbo].[@XASCH] where docnum=" & docnum
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                dr.Read()
                If (TxtSapNO.Text = "NA") Then
                    sapno = 0
                Else
                    sapno = CLng(TxtSapNO.Text)
                End If
                If (dr(0) <> TxtSubject.Text Or dr(1) <> sapno) Then
                    CommUtil.ShowMsg(Me, "資料有變動,需先儲存")
                    dr.Close()
                    connsap.Close()
                    Exit Sub
                End If
            End If
            dr.Close()
            connsap.Close()
            SqlCmd = "Select T0.sftype,T1.status from [dbo].[@XSFTT] T0 INNER JOIN [dbo].[@XASCH] T1 On T0.sfid=T1.sfid " & 'XSFTT 簽核表單種類
                "where T1.docnum=" & docnum
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            sftype = dr(0)
            status = dr(1)
            dr.Close()
            connsap.Close()
            If (status <> "R" And status <> "B") Then
                If (sftype = 1) Then
                    If (sfid <> 100 And sfid <> 101) Then '
                        SqlCmd = "select count(*) from [dbo].[@XSPMT] T0 where T0.sfid=" & sfid  '@XSPMT 簽核人員主檔
                        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap1)
                        drsap.Read()
                        If (drsap(0) = 0) Then
                            CommUtil.ShowMsg(Me, "無外部簽核人設定,請洽MIS設定")
                            drsap.Close()
                            connsap1.Close()
                            Exit Sub
                        End If
                        drsap.Close()
                        connsap1.Close()
                        'sftype=1 一定要由下述判斷歸檔人 , 因允許設多個歸檔人, 但只有一人條件符合
                        ownid = ArchivePersonCheck(sftype)
                        If (ownid = "") Then
                            Exit Sub
                        End If
                        stoponlevel = GenSinoffProcessByLevelToCT() '原本把ron設為外部簽核 , 但還是改為內部簽核(內定簽核人員還是可放上ron , 在執行內部簽核人
                        GenSinoffProcessByFormTypeToCT(stoponlevel) '原入簽核表時 , 若發現ByLevel已放入,則ByFormType就不會重覆放入)
                        GenSinoffProcessOfOthersToCT(sftype, ownid)
                        LetRestCTFormRecordToBeInformed()
                    Else '加簽單 , 簽核人員複製母單
                        'CopySignPersonFromMainDoc(CLng(TxtAttaDoc.Text))
                        If (sfid = 100) Then
                            CopySignPersonFromMainDocToCT(CLng(TxtAttaDoc.Text))
                            LetRestCTFormRecordToBeInformed()
                        ElseIf (sfid = 101) Then
                            SqlCmd = "select uid,uname,signprice,upos,seq,emailadd,signprop,innerloop from [dbo].[@XSPWT] " &
                                    "where docentry=" & CLng(TxtAttaDoc.Text) & " and seq=1"
                            drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap) '123456
                            If (drsap.HasRows) Then
                                drsap.Read()
                                If (drsap(0) <> sid_create) Then
                                    If (ResignedCheck(drsap(0), 0) = False) Then '在職員工
                                        CType(CT.FindControl("ddl_user_" & 2), DropDownList).SelectedValue = drsap(0) & " " & drsap(1) & " " & drsap(3)
                                        CT.Rows(2).Cells(1).Text = drsap(1)
                                        CT.Rows(2).Cells(2).Text = drsap(3)
                                        CType(CT.FindControl("rbl_prop_" & 2), RadioButtonList).SelectedValue = 0
                                        CType(CT.FindControl("txt_seq_" & 2), TextBox).Text = 1
                                        CT.Rows(2).Cells(5).Text = 1
                                        CT.Rows(2).Cells(6).Text = "原出借料之人"
                                    End If
                                Else
                                    CT.Rows(0).Cells(0).Text = "可選擇保留當初之領料人,刪除其他簽核人"
                                    CopySignPersonFromMainDocToCT(CLng(TxtAttaDoc.Text))
                                End If
                            End If
                            drsap.Close()
                            connsap.Close()
                        End If
                    End If
                ElseIf (sftype = 2) Then 'By部門簽核後+自訂式 , 故將核簽人員設定表 CT 顯示 , 之後走 BtnAssignSend
                    stoponlevel = GenSinoffProcessByLevelToCT()
                ElseIf (sftype = 3) Then '自訂式 , 故將核簽人員設定表 CT 顯示 , 之後走 BtnAssignSend
                    'If (sfid = 101) Then '如是返還單 , 將原單之發送者列為簽核者
                    '    SqlCmd = "select uid,uname,signprice,upos,seq,emailadd,signprop,innerloop from [dbo].[@XSPWT] " &
                    '    "where docentry=" & CLng(TxtAttaDoc.Text) & " and seq=1"
                    '    drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap) '123456
                    '    If (drsap.HasRows) Then
                    '        drsap.Read()
                    '        If (drsap(0) <> sid_create) Then
                    '            If (ResignedCheck(drsap(0), 0) = False) Then '在職員工
                    '                CType(CT.FindControl("ddl_user_" & 2), DropDownList).SelectedValue = drsap(0) & " " & drsap(1) & " " & drsap(3)
                    '                CT.Rows(2).Cells(1).Text = drsap(1)
                    '                CT.Rows(2).Cells(2).Text = drsap(3)
                    '                CType(CT.FindControl("rbl_prop_" & 2), RadioButtonList).SelectedValue = 0
                    '                CType(CT.FindControl("txt_seq_" & 2), TextBox).Text = 1
                    '                CT.Rows(2).Cells(5).Text = 1
                    '                CT.Rows(2).Cells(6).Text = "原出借料之人"
                    '            End If
                    '        End If
                    '    End If
                    '    drsap.Close()
                    '    connsap.Close()
                    'End If
                End If
                CT.Visible = True
                HT.Visible = False
                HeadT.Visible = False
                CommT.Visible = False
                FT_m.Visible = False
                FT_0.Visible = False
                FT_1.Visible = False
                iframeContent.Visible = False
                ItemT.Visible = False
                SignT.Visible = False
                ContentT.Visible = False
                AddT.Visible = False
                FormLogoTitleT.Visible = False
            Else '如果是 R或 B 因簽核人員已設定好 , 故不需再產生 , 直接修改相關欄位
                SqlCmd = "update [dbo].[@XSPWT] set comment='" & comment & "',status=100,signdate='" & Format(Now(), "yyyy/MM/dd HH:mm:ss") & "' where signprop=0 and docentry=" & docnum & " and seq=1"
                CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                connsap.Close()
                SqlCmd = "update [dbo].[@XSPWT] set status=1,receivedate='" & Format(Now(), "yyyy/MM/dd HH:mm:ss") & "' where signprop=0 And docentry=" & docnum & " And seq=2"
                CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                connsap.Close()
                historystatus = "重啟流程"
                SqlCmd = "update [dbo].[@XASCH] Set status='O' " &
                    "where docnum=" & docnum
                CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                connsap.Close()
                docstatus = "O"
                '在@XSPHT(簽核歷史Table)記錄此送審資料
                RecordSignFlowHistoty(comment, historystatus)
                'send email to first people
                Email_SignOffFlow("send", 0)
                If (actmode = "signoff" Or actmode = "recycle" Or actmode = "recycle_login" Or actmode = "signoff_login") Then
                    ds.Tables(0).Rows(Session("startindex"))("signoffflag") = 1 '設定該員簽核表單list中 , 此單設為已簽
                    Session("ds") = ds
                    signfinish = FindNextSignDoc()
                    status = ds.Tables(0).Rows(Session("startindex"))("status")
                    If (signfinish = False) Then
                        Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&actmode=recycle&status=" & status &
                                        "&docnum=" & ds.Tables(0).Rows(Session("startindex"))("docnum") &
                                        "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & ds.Tables(0).Rows(Session("startindex"))("sfid") &
                                        "&agnid=" & agnidG & "&signflowmode=" & signflowmode)
                    Else
                        If (actmode = "recycle_login" Or actmode = "signoff_login") Then
                            Response.Redirect("~/signoff/signoff.aspx?smid=sg&smode=1&act=signfinish&signflowmode=" & signflowmode)
                        Else
                            Response.Redirect("~/usermgm/logout.aspx?act=signfinish")
                        End If
                    End If
                Else
                    Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&status=" & docstatus & "&docnum=" & docnum &
                      "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & sfid & "&agnid=" & agnidG & "&signflowmode=" & signflowmode)
                End If
            End If
        End If
    End Sub
    'Protected Sub BtnSend_ClickOrg(sender As Object, e As EventArgs)
    '    '修改狀態為O
    '    '產生簽核流程表
    '    '寫入送審歷史記錄
    '    Dim sftype As Integer
    '    Dim comment, status, historystatus As String
    '    Dim stoponlevel As Boolean
    '    Dim sapno As Long
    '    Dim beapproved, signfinish As Boolean
    '    Dim maxseq As Integer
    '    Dim ownid As String
    '    ownid = ""
    '    beapproved = False
    '    historystatus = ""
    '    '先check 是否此單已被覆核過
    '    SqlCmd = "Select status " &
    '             "FROM dbo.[@XASCH] where docnum=" & docnum
    '    dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
    '    If (dr.HasRows) Then
    '        dr.Read()
    '        If (dr(0) <> docstatus) Then
    '            beapproved = True
    '            CommUtil.ShowMsg(Me, "此單已被他人覆核過,不予處理,進入下一單")
    '        End If
    '    Else
    '        CommUtil.ShowMsg(Me, Session("s_id") & " " & ds.Tables(0).Rows(Session("startindex"))("docnum") & "簽核列表中在資料庫中找不到")
    '    End If
    '    dr.Close()
    '    connsap.Close()
    '    '''''''''''''''''end
    '    If (beapproved = False) Then
    '        comment = TxtComm.Text
    '        SqlCmd = "Select subject,sapno,price " &
    '            "from [dbo].[@XASCH] where docnum=" & docnum
    '        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
    '        If (dr.HasRows) Then
    '            dr.Read()
    '            If (TxtSapNO.Text = "NA") Then
    '                sapno = 0
    '            Else
    '                sapno = CLng(TxtSapNO.Text)
    '            End If
    '            If (dr(0) <> TxtSubject.Text Or dr(1) <> sapno) Then
    '                CommUtil.ShowMsg(Me, "資料有變動,需先儲存")
    '                dr.Close()
    '                connsap.Close()
    '                Exit Sub
    '            End If
    '        End If
    '        dr.Close()
    '        connsap.Close()
    '        SqlCmd = "Select T0.sftype,T1.status from [dbo].[@XSFTT] T0 INNER JOIN [dbo].[@XASCH] T1 On T0.sfid=T1.sfid " & 'XSFTT 簽核表單種類
    '            "where T1.docnum=" & docnum
    '        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
    '        dr.Read()
    '        sftype = dr(0)
    '        status = dr(1)
    '        dr.Close()
    '        connsap.Close()
    '        If (status <> "R" And status <> "B") Then
    '            If (sftype = 1) Then
    '                If (sfid <> 100) Then '
    '                    SqlCmd = "select count(*) from [dbo].[@XSPMT] T0 where T0.sfid=" & sfid  '@XSPMT 簽核人員主檔
    '                    drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap1)
    '                    drsap.Read()
    '                    If (drsap(0) = 0) Then
    '                        CommUtil.ShowMsg(Me, "無外部簽核人設定,請洽MIS設定")
    '                        drsap.Close()
    '                        connsap1.Close()
    '                        Exit Sub
    '                    End If
    '                    drsap.Close()
    '                    connsap1.Close()
    '                    ownid = ArchivePersonCheck(sftype)
    '                    If (ownid = "") Then
    '                        Exit Sub
    '                    End If
    '                    'seq=1 之 signdate , status 在 GenSinoffProcessByLevel 之InsertSignOffProcessRecord 寫入
    '                    'seq=2 之 status 在 GenSinoffProcessByLevel 之InsertSignOffProcessRecord 寫入
    '                    'stoponlevel = GenSinoffProcessByLevel() '原本把ron設為外部簽核 , 但還是改為內部簽核(內定簽核人員還是可放上ron , 在執行內部簽核人
    '                    'GenSinoffProcessByFormType(stoponlevel) '原入簽核表時 , 若發現ByLevel已放入,則ByFormType就不會重覆放入)
    '                    'GenSinoffProcessOfOthers(sftype, ownid)
    '                    stoponlevel = GenSinoffProcessByLevelToCT() '原本把ron設為外部簽核 , 但還是改為內部簽核(內定簽核人員還是可放上ron , 在執行內部簽核人
    '                    GenSinoffProcessByFormTypeToCT(stoponlevel) '原入簽核表時 , 若發現ByLevel已放入,則ByFormType就不會重覆放入)
    '                    GenSinoffProcessOfOthersToCT(sftype, ownid)
    '                Else '加簽單 , 簽核人員複製母單
    '                    CopySignPersonFromMainDoc(CLng(TxtAttaDoc.Text))
    '                End If
    '                'update seq=1 之comment (signdate 及status 已在GenSinoffProcessByLevel()中之Insert 發送者時產生)
    '                SqlCmd = "update [dbo].[@XSPWT] Set comment='" & comment & "' where signprop=0 and docentry=" & docnum & " and seq=1"
    '                CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
    '                connsap.Close()

    '                SqlCmd = "Select max(seq) from [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum
    '                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
    '                dr.Read()
    '                maxseq = dr(0)
    '                dr.Close()
    '                connsap.Close()
    '                If (maxseq <> 1) Then '處理只有一個人,不需審合情況 , 如領用低值物品或其他......
    '                    'update seq=2 之receivedate
    '                    SqlCmd = "update [dbo].[@XSPWT] Set receivedate='" & Format(Now(), "yyyy/MM/dd HH:mm:ss") & "' where signprop=0 And docentry=" & docnum & " And seq=2"
    '                    CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
    '                    connsap.Close()
    '                    historystatus = "啟動流程"
    '                    SqlCmd = "update [dbo].[@XASCH] Set status='O' " &
    '                        "where docnum=" & docnum
    '                    CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
    '                    connsap.Close()
    '                    docstatus = "O"
    '                Else
    '                    SqlCmd = "update [dbo].[@XASCH] set status='F' where docnum=" & docnum
    '                    CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
    '                    connsap.Close()
    '                    historystatus = "啟動流程並結案"
    '                    docstatus = "F"
    '                    '設歸檔人員 status=1及 receivedate
    '                    SqlCmd = "update [dbo].[@XSPWT] set status=1,receivedate='" & Format(Now(), "yyyy/MM/dd HH:mm:ss") & "' where signprop=1 and docentry=" & docnum & " And seq=2"
    '                    CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
    '                    connsap.Close()
    '                    '設知悉人員 receivedate
    '                    SqlCmd = "update [dbo].[@XSPWT] set status=1,receivedate='" & Format(Now(), "yyyy/MM/dd HH:mm:ss") & "' where signprop=2 and docentry=" & docnum
    '                    CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
    '                    connsap.Close()
    '                End If
    '            ElseIf (sftype = 3) Then '自訂式 , 故將核簽人員設定表 CT 顯示 , 之後走 BtnAssignSend
    '                CT.Visible = True
    '                HT.Visible = False
    '                HeadT.Visible = False
    '                CommT.Visible = False
    '                FT_m.Visible = False
    '                FT_0.Visible = False
    '                FT_1.Visible = False
    '                iframeContent.Visible = False
    '                ItemT.Visible = False
    '                SignT.Visible = False
    '                ContentT.Visible = False
    '                FormLogoTitleT.Visible = False
    '                'CreateSignFlowPerson()
    '                Exit Sub
    '            End If
    '        Else '如果是 R或 B 因簽核人員已設定好 , 故不需再產生 , 直接修改相關欄位
    '            SqlCmd = "update [dbo].[@XSPWT] set comment='" & comment & "',status=100,signdate='" & Format(Now(), "yyyy/MM/dd HH:mm:ss") & "' where signprop=0 and docentry=" & docnum & " and seq=1"
    '            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
    '            connsap.Close()
    '            SqlCmd = "update [dbo].[@XSPWT] set status=1,receivedate='" & Format(Now(), "yyyy/MM/dd HH:mm:ss") & "' where signprop=0 And docentry=" & docnum & " And seq=2"
    '            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
    '            connsap.Close()
    '            historystatus = "重啟流程"
    '            SqlCmd = "update [dbo].[@XASCH] Set status='O' " &
    '                "where docnum=" & docnum
    '            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
    '            connsap.Close()
    '            docstatus = "O"
    '        End If

    '        '在@XSPHT(簽核歷史Table)記錄此送審資料
    '        RecordSignFlowHistoty(comment, historystatus)
    '        'send email to first people
    '        Email_SignOffFlow("send", 0)
    '    End If
    '    If (actmode = "signoff" Or actmode = "recycle" Or actmode = "recycle_login" Or actmode = "signoff_login") Then
    '        ds.Tables(0).Rows(Session("startindex"))("signoffflag") = 1 '設定該員簽核表單list中 , 此單設為已簽
    '        Session("ds") = ds
    '        signfinish = FindNextSignDoc()
    '        status = ds.Tables(0).Rows(Session("startindex"))("status")
    '        If (signfinish = False) Then
    '            Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&actmode=recycle&status=" & status &
    '                            "&docnum=" & ds.Tables(0).Rows(Session("startindex"))("docnum") &
    '                            "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & ds.Tables(0).Rows(Session("startindex"))("sfid") &
    '                            "&agnid=" & agnidG & "&signflowmode=" & signflowmode)
    '        Else
    '            If (actmode = "recycle_login" Or actmode = "signoff_login") Then
    '                Response.Redirect("~/signoff/signoff.aspx?smid=sg&smode=1&act=signfinish&signflowmode=" & signflowmode)
    '            Else
    '                Response.Redirect("~/usermgm/logout.aspx?act=signfinish")
    '            End If
    '        End If
    '    Else
    '        Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&status=" & docstatus & "&docnum=" & docnum &
    '          "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & sfid & "&agnid=" & agnidG & "&signflowmode=" & signflowmode)
    '    End If
    'End Sub
    Sub CreateSignFlowPerson()
        Dim i As Integer
        Dim tRow As TableRow
        Dim tCell As TableCell
        Dim RBLProp As RadioButtonList
        Dim TxtID As TextBox
        Dim idstr As String
        'Dim DDLUser As DropDownExtender
        Dim DDLUser As DropDownList
        Dim Labelx As Label
        Dim sftype As Integer
        SqlCmd = "Select T0.sftype from [dbo].[@XSFTT] T0 " & 'XSFTT 簽核表單種類
                "where T0.sfid=" & sfid
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        dr.Read()
        sftype = dr(0)
        dr.Close()
        connsap.Close()
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.BackColor = Drawing.Color.LightGreen
        tRow.Font.Bold = True
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 7
        tCell.HorizontalAlign = HorizontalAlign.Center
        If (sftype = 3) Then
            'If (sfid <> 101) Then
            DDLSignDefault = New DropDownList
                DDLSignDefault.ID = "ddl_signdefault"
                DDLSignDefault.Width = 180
                DDLSignDefault.AutoPostBack = True
                SqlCmd = "Select distinct T0.signpid,T0.signpname from dbo.[@XSPAT] T0 where ownid='" & Session("s_id") & "'"
                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                DDLSignDefault.Items.Clear()
                DDLSignDefault.Items.Add("可選擇預設簽核群組")
                If (dr.HasRows) Then
                    Do While (dr.Read())
                        DDLSignDefault.Items.Add(dr(1) & " " & dr(0))
                    Loop
                End If
                dr.Close()
                connsap.Close()
                AddHandler DDLSignDefault.SelectedIndexChanged, AddressOf DDLSignDefault_SelectedIndexChanged
                tCell.Controls.Add(DDLSignDefault)
            'Else
            'tCell.Text = "選擇出借人及需要之知悉人"
            'End If
        ElseIf (sftype = 2) Then
            tCell.Text = "簽核人內含部門主管(但可再新增其他人)"
        Else
            If (sfid <> 101) Then
                tCell.Text = "簽核人已內定(但可再新增知悉人)"
            Else
                tCell.Text = "選擇出借人及需要之知悉人"
            End If
        End If
        'tCell.Text = "簽核人員設定"
        tRow.Controls.Add(tCell)
        CT.Rows.Add(tRow)

        tRow = New TableRow()
        tRow.BackColor = Drawing.Color.LightGreen
        tRow.Font.Bold = True
        For i = 1 To 7
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
                tCell.Text = "在職"
            ElseIf (i = 7) Then
                tCell.Text = "備註"
            End If
            tRow.Controls.Add(tCell)
        Next
        CT.Rows.Add(tRow)
        For i = 2 To signpersonmaxrow + 1
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
            DDLUser.Width = 80
            AddHandler DDLUser.SelectedIndexChanged, AddressOf DDLUser_SelectedIndexChanged
            DDLUser.AutoPostBack = True
            tCell.Controls.Add(DDLUser)
            DDLUser.Items.Clear()
            DDLUser.Items.Add("")
            'SqlCmd = "select id,name,position from dbo.[user] where denyf<>1 and email<>'' and id <> '" & Session("s_id") & "' order by branch,grp"
            SqlCmd = "select id,name,position from dbo.[user] where denyf<>1 and email<>'' order by branch,grp"
            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn) 'modify
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
            'If (sfid = 100 Or sfid = 101) Then
            '    RBLProp.Items.Add("借料人")
            'End If
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

            tCell = New TableCell '在職
            tCell.BorderWidth = 1
            tCell.Wrap = False
            tCell.HorizontalAlign = HorizontalAlign.Center
            tCell.Width = 40
            tRow.Controls.Add(tCell)

            tCell = New TableCell '備註
            tCell.BorderWidth = 1
            tCell.Wrap = False
            tCell.HorizontalAlign = HorizontalAlign.Center
            tCell.Width = 200
            tRow.Controls.Add(tCell)

            CT.Rows.Add(tRow)
        Next
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.BackColor = Drawing.Color.LightGreen
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 7
        tCell.HorizontalAlign = HorizontalAlign.Center
        BtnAssignSend = New Button
        BtnAssignSend.ID = "btn_assignsend"
        BtnAssignSend.Width = 40
        BtnAssignSend.Text = "送審"
        'BtnAssignSend.OnClientClick = "Return confirm('要送審嗎')"
        AddHandler BtnAssignSend.Click, AddressOf BtnAssignSend_Click
        tCell.Controls.Add(BtnAssignSend)
        Labelx = New Label()
        Labelx.ID = "label_assignsend"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        BtnAssignCancel = New Button
        BtnAssignCancel.ID = "btn_assigncancel"
        BtnAssignCancel.Width = 40
        BtnAssignCancel.Text = "取消"
        'BtnAssignCancel.OnClientClick = "Return confirm('要取消嗎')"
        AddHandler BtnAssignCancel.Click, AddressOf BtnAssignCancel_Click
        tCell.Controls.Add(BtnAssignCancel)
        tRow.Controls.Add(tCell)
        CT.Rows.Add(tRow)
    End Sub
    Protected Sub DDLSignDefault_SelectedIndexChanged(sender As Object, e As EventArgs) 'ron
        Dim sfid, i As Integer
        i = 2
        If (DDLSignDefault.SelectedIndex = 0) Then ''可以選擇用 Redirect 方式 , 也可如下直接對Table處理
            'Response.Redirect("~/signoff/signoffsetup.aspx?smid=sg&smode=3&mode=ctinit")
            ClearItemsOfSignOffPerson()
        Else '可以選擇用 Redirect 方式(但此方式要注意一些參數之延續性) , 也可如下直接對Table處理
            Dim str() As String
            ClearItemsOfSignOffPerson()
            str = Split(DDLSignDefault.SelectedValue, " ")
            sfid = str(1)
            SqlCmd = "select T0.uid,T0.seq,T0.prop,T0.num,T0.signpname from [@XSPAT] T0 where T0.process=0 and T0.sgtype=0 and T0.signpid=" & sfid &
                    " order by T0.seq"
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                Do While (dr.Read())
                    SqlCmd = "select T0.name,T0.position,T0.denyf from dbo.[User] T0 where T0.id='" & dr(0) & "'"
                    dr1 = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                    dr1.Read()
                    If (dr1(2) = 1) Then
                        CType(CT.FindControl("ddl_user_" & i), DropDownList).Items.Add(dr(0) & " " & dr1(0) & " " & dr1(1))
                    End If
                    CType(CT.FindControl("ddl_user_" & i), DropDownList).SelectedValue = dr(0) & " " & dr1(0) & " " & dr1(1)
                    CT.Rows(i).Cells(1).Text = dr1(0)
                    CT.Rows(i).Cells(2).Text = dr1(1)
                    If (dr1(2) = 0) Then
                        CT.Rows(i).Cells(5).Text = 1
                    Else
                        CT.Rows(i).Cells(5).Text = 0
                        CT.Rows(i).Cells(6).Text = "離職狀態,不會列入簽核"
                    End If
                    CType(CT.FindControl("rbl_prop_" & i), RadioButtonList).SelectedValue = dr(2)
                    CType(CT.FindControl("rbl_prop_" & i), RadioButtonList).Enabled = True
                    CType(CT.FindControl("txt_seq_" & i), TextBox).Text = dr(1)
                    If (dr(1) <> 0) Then
                        CType(CT.FindControl("txt_seq_" & i), TextBox).Enabled = True
                    Else
                        CType(CT.FindControl("txt_seq_" & i), TextBox).Enabled = False
                    End If
                    dr1.Close()
                    conn.Close()
                    i = i + 1
                Loop
            End If
            dr.Close()
            connsap.Close()
            'CT.Visible = True
            'Response.Redirect("~/signoff/signoffsetup.aspx?smid=sg&smode=3&mode=ctimport&ddlsigndefaultindex=" & DDLSignDefault.SelectedIndex & "&sfid=" & sfid)
            'PutDataToForm()
        End If
    End Sub
    Protected Sub DDLUser_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim idstr As String
        Dim str() As String
        Dim row As Integer
        Dim sftype As Integer
        Dim i As Integer
        i = 2
        SqlCmd = "Select T0.sftype from [dbo].[@XSFTT] T0 " & 'XSFTT 簽核表單種類
                "where T0.sfid=" & sfid
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        dr.Read()
        sftype = dr(0)
        dr.Close()
        connsap.Close()
        idstr = sender.ID
        str = Split(idstr, "_")
        row = CInt(str(2))
        CT.Rows(row).Cells(6).Text = ""
        If (sender.SelectedIndex <> 0) Then
            str = Split(sender.SelectedValue, " ")
            If (str(0) <> Session("s_id")) Then
                If (sftype = 3 Or sftype = 2) Then
                    CType(CT.FindControl("ddl_user_" & row), DropDownList).SelectedValue = str(0) & " " & str(1) & " " & str(2)
                    CType(CT.FindControl("rbl_prop_" & row), RadioButtonList).Enabled = True
                    CType(CT.FindControl("txt_seq_" & row), TextBox).Text = row - 1
                    CT.Rows(row).Cells(1).Text = str(1)
                    CT.Rows(row).Cells(2).Text = str(2)
                    CT.Rows(row).Cells(5).Text = 1
                ElseIf (sftype = 1) Then
                    For i = 2 To row - 1
                        If (CType(CT.FindControl("ddl_user_" & i), DropDownList).SelectedValue = sender.SelectedValue) Then
                            CommUtil.ShowMsg(Me, "選擇的人員 " & sender.SelectedValue & " 已存在")
                            CType(CT.FindControl("ddl_user_" & row), DropDownList).SelectedValue = ""
                            Exit Sub
                        End If
                    Next
                    CType(CT.FindControl("ddl_user_" & row), DropDownList).SelectedValue = str(0) & " " & str(1) & " " & str(2)
                    CType(CT.FindControl("rbl_prop_" & row), RadioButtonList).SelectedValue = 2
                    CType(CT.FindControl("txt_seq_" & row), TextBox).Text = 0
                    CT.Rows(row).Cells(1).Text = str(1)
                    CT.Rows(row).Cells(2).Text = str(2)
                    CT.Rows(row).Cells(5).Text = 1
                Else
                    CommUtil.ShowMsg(Me, "無表單種類" & sftype & "之設計")
                    CType(CT.FindControl("ddl_user_" & row), DropDownList).SelectedValue = ""
                End If
            Else
                CommUtil.ShowMsg(Me, "不能選擇自己")
                CType(CT.FindControl("ddl_user_" & row), DropDownList).SelectedValue = ""
            End If
        Else
            CType(CT.FindControl("ddl_user_" & row), DropDownList).SelectedValue = ""
            CType(CT.FindControl("txt_seq_" & row), TextBox).Text = ""
            CType(CT.FindControl("rbl_prop_" & row), RadioButtonList).SelectedIndex = -1
            CType(CT.FindControl("rbl_prop_" & row), RadioButtonList).Enabled = False
            CT.Rows(row).Cells(1).Text = ""
            CT.Rows(row).Cells(2).Text = ""
            CT.Rows(row).Cells(5).Text = ""

        End If
    End Sub
    Protected Sub DDLAttaFile_SelectedIndexChanged(sender As Object, e As EventArgs)
        ShowIframeContent()
    End Sub
    Sub ShowIframeContent()
        If (DDLAttaFile.SelectedIndex <> 0) Then
            Dim httpfile As String
            Dim siddir As String
            Dim sfidnum As Integer
            Dim str() As String
            str = Split(DDLAttaFile.SelectedValue, "_")
            SqlCmd = "Select sid,sfid from [dbo].[@XASCH] where docnum=" & CLng(str(0))
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            siddir = dr(0)
            sfidnum = dr(1)
            dr.Close()
            connsap.Close()
            httpfile = url & "AttachFile/SignOffsFormFiles/" & siddir & "/" & sfidnum & "/" & DDLAttaFile.SelectedValue
            iframeContent.Visible = True
            iframeContent.Attributes.Remove("src")
            iframeContent.Attributes.Add("src", httpfile)
            If (docstatus = "D" Or docstatus = "A" Or docstatus = "E" Or ((docstatus = "R" Or docstatus = "B") And formstatusindex = 0)) Then
                CType(FT_m.FindControl("chk_del_m"), CheckBox).Visible = True
            End If
        Else
            iframeContent.Attributes.Remove("src")
            CType(FT_m.FindControl("chk_del_m"), CheckBox).Visible = False
            iframeContent.Visible = False
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
    End Sub

    Function SignOffFlowPersonFieldCheck()
        Dim i, j As Integer
        Dim uid As String
        Dim prop, propc, proprow, targetseq, count As Integer
        Dim str() As String
        SignOffFlowPersonFieldCheck = True
        count = 0
        For i = 2 To signpersonmaxrow + 1
            str = Split(CType(CT.FindControl("ddl_user_" & i), DropDownList).SelectedValue, " ")
            uid = str(0)
            If (uid <> "") Then
                If (CType(CT.FindControl("rbl_prop_" & i), RadioButtonList).SelectedIndex = -1) Then
                    CommUtil.ShowMsg(Me, uid & "(" & CT.Rows(i).Cells(1).Text & ") 屬性沒設定")
                    SignOffFlowPersonFieldCheck = False
                    Exit Function
                End If
                'MsgBox(CType(CT.FindControl("rbl_prop_" & i), RadioButtonList).SelectedValue)
                prop = CType(CT.FindControl("rbl_prop_" & i), RadioButtonList).SelectedValue
                If (prop = 1) Then
                    propc = propc + 1
                    proprow = i
                End If
                If (Not IsNumeric(CType(CT.FindControl("txt_seq_" & i), TextBox).Text)) Then
                    CommUtil.ShowMsg(Me, uid & "(" & CT.Rows(i).Cells(1).Text & ") 順序沒設定或設定的不是數值")
                    SignOffFlowPersonFieldCheck = False
                    Exit Function
                End If
                targetseq = CInt(CType(CT.FindControl("txt_seq_" & i), TextBox).Text)
                If (targetseq <> 0) Then
                    For j = i + 1 To signpersonmaxrow + 1
                        If (CType(CT.FindControl("txt_seq_" & j), TextBox).Text <> "") Then
                            If (targetseq = CInt(CType(CT.FindControl("txt_seq_" & j), TextBox).Text)) Then
                                CommUtil.ShowMsg(Me, "順序" & targetseq & "有重複 , 請修正")
                                SignOffFlowPersonFieldCheck = False
                                Exit Function
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
                                SignOffFlowPersonFieldCheck = False
                                Exit Function
                            End If
                        End If
                    End If
                Next
                count = count + 1
            End If
        Next
        If (count = 0) Then
            CommUtil.ShowMsg(Me, "簽核表單是空的,最少要一個")
            SignOffFlowPersonFieldCheck = False
            Exit Function
        End If
        If (propc >= 2) Then
            CommUtil.ShowMsg(Me, "簽核表單歸檔設定只能有1個")
            SignOffFlowPersonFieldCheck = False
            Exit Function
        End If
    End Function
    Function ResignedCheck(uid As String, sftype As Integer)
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        ResignedCheck = False
        SqlCmd = "select denyf,name from dbo.[user] where id='" & uid & "'"
        drL = CommUtil.SelectLocalSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            drL.Read()
            If (drL(0) = 1) Then
                If (sftype = 1 Or sftype = 2) Then
                    CommUtil.ShowMsg(Me, "發現帳號:" & uid & "(" & drL(1) & ") 為離職員工,出現在內建簽核表中,請洽管理員修正")
                ElseIf (sftype = 3) Then
                    CommUtil.ShowMsg(Me, "發現帳號:" & uid & "(" & drL(1) & ") 為離職員工,出現在自訂簽核表中,請自行修正")
                Else
                    'nothing
                End If
                ResignedCheck = True
            End If
        End If
        drL.Close()
        connL.Close()
    End Function

    Function SignOffPersonSetSave()
        Dim i, seq, prop As Integer
        Dim uid As String
        Dim str() As String
        Dim ownid As String
        ownid = ""
        'del 
        SqlCmd = "delete from [dbo].[@XSPAT] where sgtype=1 and ownid='" & sid_create & "'"
        CommUtil.SqlSapExecute("del", SqlCmd, connsap)
        connsap.Close()

        For i = 2 To signpersonmaxrow + 1
            If (CT.Rows(i).Cells(5).Text = "1") Then
                str = Split(CType(CT.FindControl("ddl_user_" & i), DropDownList).SelectedValue, " ")
                uid = str(0)
                If (uid <> "") Then
                    seq = CInt(CType(CT.FindControl("txt_seq_" & i), TextBox).Text)
                    prop = CType(CT.FindControl("rbl_prop_" & i), RadioButtonList).SelectedValue
                    If (prop = 1) Then
                        ownid = uid
                    End If
                    SqlCmd = "insert into [dbo].[@XSPAT] (uid,seq,prop,ownid,sgtype) " &
                             "values('" & uid & "'," & seq & "," & prop & ",'" & sid_create & "',1)"
                    CommUtil.SqlSapExecute("ins", SqlCmd, connsap)
                    connsap.Close()
                End If
            End If
        Next
        '以seq排序讀出此設定, 再以1開始,重新順序寫入seq
        SqlCmd = "Select T0.seq from [@XSPAT] T0 where T0.sgtype=1 and T0.prop=0 And T0.ownid='" & sid_create & "' order by T0.seq"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            i = 1
            Do While (dr.Read())
                If (i <> dr(0)) Then
                    SqlCmd = "Update [dbo].[@XSPAT] set seq=" & i & " " &
                        "where sgtype=1 and seq=" & dr(0) & " and ownid='" & sid_create & "'"
                    CommUtil.SqlSapExecute("upd", SqlCmd, conn)
                    conn.Close()
                End If
                i = i + 1
            Loop
        End If
        dr.Close()
        connsap.Close()
        If (ownid = "") Then
            ownid = Session("s_id")
        End If
        Return ownid
    End Function
    Sub ClearItemsOfSignOffPerson()
        Dim i As Integer
        Dim uid As String
        For i = 2 To signpersonmaxrow + 1
            uid = CType(CT.FindControl("ddl_user_" & i), DropDownList).SelectedValue
            If (uid <> "") Then
                CType(CT.FindControl("ddl_user_" & i), DropDownList).SelectedIndex = 0
                CType(CT.FindControl("txt_seq_" & i), TextBox).Text = ""
                CType(CT.FindControl("rbl_prop_" & i), RadioButtonList).SelectedIndex = -1
                CT.Rows(i).Cells(1).Text = ""
                CT.Rows(i).Cells(2).Text = ""
                CT.Rows(i).Cells(5).Text = ""
            End If
        Next
    End Sub
    Protected Sub BtnAssignSend_Click(sender As Object, e As EventArgs)
        Dim sftype As Integer
        Dim comment, status, historystatus As String
        Dim signfinish As Boolean
        Dim maxseq As Integer
        Dim ownid As String
        'Dim sapno As Long
        ownid = ""
        'comment = TxtComm.Text
        'SqlCmd = "Select subject,sapno,price " &
        '    "from [dbo].[@XASCH] where docnum=" & docnum
        'dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        'If (dr.HasRows) Then
        '    dr.Read()
        '    If (TxtSapNO.Text = "NA") Then
        '        sapno = 0
        '    Else
        '        sapno = CLng(TxtSapNO.Text)
        '    End If
        '    If (dr(0) <> TxtSubject.Text Or dr(1) <> sapno) Then
        '        CommUtil.ShowMsg(Me, "資料有變動,需先儲存")
        '        dr.Close()
        '        connsap.Close()
        '        Exit Sub
        '    End If
        'End If
        'dr.Close()
        'connsap.Close()
        SqlCmd = "Select T0.sftype,T1.status from [dbo].[@XSFTT] T0 INNER JOIN [dbo].[@XASCH] T1 On T0.sfid=T1.sfid " & 'XSFTT 簽核表單種類
            "where T1.docnum=" & docnum
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        dr.Read()
        sftype = dr(0)
        status = dr(1)
        dr.Close()
        connsap.Close()
        'check 簽核人員設定表
        If (sftype = 3) Then
            If (SignOffFlowPersonFieldCheck() = False) Then
                Exit Sub
            End If
        End If
        ownid = SignOffPersonSetSave()

        '寫入簽核表 XSPWT
        '修改狀態為O
        '產生簽核流程表
        '寫入送審歷史記錄
        GenSinoffProcessRecordsOfMain()
        GenSinoffProcessRecordsOfOthers(ownid)
        comment = TxtComm.Text
        'update seq=1 之comment (signdate 及status 已在GenSinoffProcessByLevel()中之Insert 發送者時產生)'kkkkkkk
        SqlCmd = "update [dbo].[@XSPWT] set comment='" & comment & "' where signprop=0 and docentry=" & docnum & " and seq=1"
        CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
        connsap.Close()

        SqlCmd = "Select max(seq) from [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        dr.Read()
        maxseq = dr(0)
        dr.Close()
        connsap.Close()
        If (maxseq <> 1) Then
            'update seq=2
            SqlCmd = "update [dbo].[@XSPWT] set receivedate='" & Format(Now(), "yyyy/MM/dd HH:mm:ss") & "' where signprop=0 And docentry=" & docnum & " And seq=2"
            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
            connsap.Close()
            historystatus = "啟動流程"

            SqlCmd = "update [dbo].[@XASCH] set status='O' " &
                "where docnum=" & docnum
            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
            connsap.Close()
            docstatus = "O"
        Else
            SqlCmd = "update [dbo].[@XASCH] set status='F' where docnum=" & docnum
            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
            connsap.Close()
            historystatus = "啟動流程並結案"
            docstatus = "F"
            '設歸檔人員 status=1及 receivedate
            SqlCmd = "update [dbo].[@XSPWT] set status=1,receivedate='" & Format(Now(), "yyyy/MM/dd HH:mm:ss") & "' where signprop=1 and docentry=" & docnum & " And seq=2"
            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
            connsap.Close()
            '設知悉人員 receivedate
            SqlCmd = "update [dbo].[@XSPWT] set status=1,receivedate='" & Format(Now(), "yyyy/MM/dd HH:mm:ss") & "' where signprop=2 and docentry=" & docnum
            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
            connsap.Close()
        End If
        '在@XSPHT(簽核歷史Table)記錄此送審資料
        RecordSignFlowHistoty(comment, historystatus)
        'send email to first people
        Email_SignOffFlow("send", 0)
        'Cleaar items
        ClearItemsOfSignOffPerson()
        '回覆原表單畫面
        CT.Visible = False
        HT.Visible = True
        HeadT.Visible = True
        CommT.Visible = True
        FT_m.Visible = True
        FT_0.Visible = True
        FT_1.Visible = True
        iframeContent.Visible = True
        ItemT.Visible = True
        SignT.Visible = True
        ContentT.Visible = True
        FormLogoTitleT.Visible = True

        If (actmode = "signoff" Or actmode = "recycle" Or actmode = "recycle_login" Or actmode = "signoff_login") Then
            ds.Tables(0).Rows(Session("startindex"))("signoffflag") = 1 '設定該員簽核表單list中 , 此單設為已簽
            Session("ds") = ds
            signfinish = FindNextSignDoc()
            status = ds.Tables(0).Rows(Session("startindex"))("status")
            If (signfinish = False) Then

                Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&actmode=recycle&status=" & status &
                                "&docnum=" & ds.Tables(0).Rows(Session("startindex"))("docnum") &
                                "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & ds.Tables(0).Rows(Session("startindex"))("sfid") &
                                "&agnid=" & agnidG & "&signflowmode=" & signflowmode)
            Else
                If (actmode = "recycle_login" Or actmode = "signoff_login") Then
                    Response.Redirect("~/signoff/signoff.aspx?smid=sg&smode=1&act=signfinish&signflowmode=" & signflowmode)
                Else
                    Response.Redirect("~/usermgm/logout.aspx?act=signfinish")
                End If
            End If
        Else
            Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&status=" & docstatus & "&docnum=" & docnum &
              "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & sfid & "&agnid=" & agnidG & "&signflowmode=" & signflowmode)
        End If

    End Sub

    'Protected Sub BtnAssignSendORG_Click(sender As Object, e As EventArgs)
    '    Dim sftype As Integer
    '    Dim comment, status, historystatus As String
    '    Dim beapproved, signfinish As Boolean
    '    Dim maxseq As Integer
    '    Dim ownid As String
    '    ownid = ""
    '    beapproved = False
    '    '先check 是否此單已被覆核過
    '    SqlCmd = "Select status " &
    '             "FROM dbo.[@XASCH] where docnum=" & docnum
    '    dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
    '    If (dr.HasRows) Then
    '        dr.Read()
    '        If (dr(0) <> docstatus) Then
    '            beapproved = True
    '            CommUtil.ShowMsg(Me, "此單已被他人覆核過,不予處理,進入下一單")
    '        End If
    '    Else
    '        CommUtil.ShowMsg(Me, Session("s_id") & " " & ds.Tables(0).Rows(Session("startindex"))("docnum") & "簽核列表中在資料庫中找不到")
    '    End If
    '    dr.Close()
    '    connsap.Close()
    '    '''''''''''''''''end
    '    If (beapproved = False) Then
    '        'check 簽核人員設定表
    '        If (SignOffFlowPersonFieldCheck() = False) Then
    '            Exit Sub
    '        Else
    '            SignOffPersonSetSave()
    '        End If

    '        '寫入簽核表 XSPWT
    '        '修改狀態為O
    '        '產生簽核流程表
    '        '寫入送審歷史記錄

    '        SqlCmd = "select T0.sftype,T1.status from [dbo].[@XSFTT] T0 INNER JOIN [dbo].[@XASCH] T1 ON T0.sfid=T1.sfid " & 'XSFTT 簽核表單種類
    '            "where T1.docnum=" & docnum
    '        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
    '        dr.Read()
    '        sftype = dr(0)
    '        status = dr(1)
    '        dr.Close()
    '        connsap.Close()
    '        ownid = ArchivePersonCheck(sftype)
    '        If (ownid = "") Then
    '            Exit Sub
    '        End If
    '        GenSinoffProcessByAssignment()
    '        GenSinoffProcessOfOthers(sftype, ownid)
    '        comment = TxtComm.Text
    '        'update seq=1 之comment (signdate 及status 已在GenSinoffProcessByLevel()中之Insert 發送者時產生)
    '        SqlCmd = "update [dbo].[@XSPWT] set comment='" & comment & "' where signprop=0 and docentry=" & docnum & " and seq=1"
    '        CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
    '        connsap.Close()

    '        SqlCmd = "Select max(seq) from [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum
    '        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
    '        dr.Read()
    '        maxseq = dr(0)
    '        dr.Close()
    '        connsap.Close()
    '        If (maxseq <> 1) Then
    '            'update seq=2
    '            SqlCmd = "update [dbo].[@XSPWT] set receivedate='" & Format(Now(), "yyyy/MM/dd HH:mm:ss") & "' where signprop=0 And docentry=" & docnum & " And seq=2"
    '            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
    '            connsap.Close()
    '            historystatus = "啟動流程"

    '            SqlCmd = "update [dbo].[@XASCH] set status='O' " &
    '                "where docnum=" & docnum
    '            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
    '            connsap.Close()
    '            docstatus = "O"
    '        Else
    '            SqlCmd = "update [dbo].[@XASCH] set status='F' where docnum=" & docnum
    '            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
    '            connsap.Close()
    '            historystatus = "啟動流程並結案"
    '            docstatus = "F"
    '            '設歸檔人員 status=1及 receivedate
    '            SqlCmd = "update [dbo].[@XSPWT] set status=1,receivedate='" & Format(Now(), "yyyy/MM/dd HH:mm:ss") & "' where signprop=1 and docentry=" & docnum & " And seq=2"
    '            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
    '            connsap.Close()
    '            '設知悉人員 receivedate
    '            SqlCmd = "update [dbo].[@XSPWT] set status=1,receivedate='" & Format(Now(), "yyyy/MM/dd HH:mm:ss") & "' where signprop=2 and docentry=" & docnum
    '            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
    '            connsap.Close()
    '        End If
    '        '在@XSPHT(簽核歷史Table)記錄此送審資料
    '        RecordSignFlowHistoty(comment, historystatus)
    '        'send email to first people
    '        Email_SignOffFlow("send", 0)
    '        'Cleaar items
    '        ClearItemsOfSignOffPerson()
    '        '回覆原表單畫面
    '        CT.Visible = False
    '        HT.Visible = True
    '        HeadT.Visible = True
    '        CommT.Visible = True
    '        FT_m.Visible = True
    '        FT_0.Visible = True
    '        FT_1.Visible = True
    '        iframeContent.Visible = True
    '        ItemT.Visible = True
    '        SignT.Visible = True
    '        ContentT.Visible = True
    '        FormLogoTitleT.Visible = True
    '    End If
    '    If (actmode = "signoff" Or actmode = "recycle" Or actmode = "recycle_login" Or actmode = "signoff_login") Then
    '        ds.Tables(0).Rows(Session("startindex"))("signoffflag") = 1 '設定該員簽核表單list中 , 此單設為已簽
    '        Session("ds") = ds
    '        signfinish = FindNextSignDoc()
    '        status = ds.Tables(0).Rows(Session("startindex"))("status")
    '        If (signfinish = False) Then

    '            Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&actmode=recycle&status=" & status &
    '                            "&docnum=" & ds.Tables(0).Rows(Session("startindex"))("docnum") &
    '                            "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & ds.Tables(0).Rows(Session("startindex"))("sfid") &
    '                            "&agnid=" & agnidG & "&signflowmode=" & signflowmode)
    '        Else
    '            If (actmode = "recycle_login" Or actmode = "signoff_login") Then
    '                Response.Redirect("~/signoff/signoff.aspx?smid=sg&smode=1&act=signfinish&signflowmode=" & signflowmode)
    '            Else
    '                Response.Redirect("~/usermgm/logout.aspx?act=signfinish")
    '            End If
    '        End If
    '    Else
    '        Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&status=" & docstatus & "&docnum=" & docnum &
    '          "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & sfid & "&agnid=" & agnidG & "&signflowmode=" & signflowmode)
    '    End If

    'End Sub
    Protected Sub BtnAssignCancel_Click(sender As Object, e As EventArgs)
        CT.Visible = False
        HT.Visible = True
        HeadT.Visible = True
        CommT.Visible = True
        FT_m.Visible = True
        FT_0.Visible = True
        FT_1.Visible = True
        iframeContent.Visible = True
        ItemT.Visible = True
        SignT.Visible = True
        If (actmode = "signoff" Or actmode = "signoff_login") Then
            Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&actmode=recycle&status=" & docstatus & "&docnum=" & docnum &
                         "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & sfid & "&agnid=" & agnidG &
                         "&signflowmode=" & signflowmode)
        Else
            Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&actmode=single&status=" & docstatus & "&docnum=" & docnum &
                         "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & sfid & "&agnid=" & agnidG &
                         "&signflowmode=" & signflowmode)
        End If
    End Sub
    Function GenSinoffProcessByLevel()
        '此為內部門流程所經人員
        '用signlevel 一直找下去 , 直到碰到id是外部審核人員或核簽金額(有金額之表單)為止
        '之後會根據每種表單所設置之流程人員 (@XSPMT 簽核人員主檔)或核簽金額(有金額之表單), 串接成完整流程
        Dim id, uname, position, emailadd As String
        Dim i As Integer
        Dim topsignoffs As Integer
        Dim signprice As Long
        Dim stoponlevel As Boolean
        Dim currate As Single
        currate = 1
        If (DDLDollorUnit.SelectedIndex <> 0) Then
            If (DDLDollorUnit.SelectedValue <> "NTD") Then
                SqlCmd = "SELECT T0.[Rate], T0.[RateDate] FROM ORTT T0 where T0.[Currency]='" & DDLDollorUnit.SelectedValue & "' order by ratedate desc"
                dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                dr.Read()
                currate = dr(0)
                dr.Close()
                conn.Close()
            End If
        End If
        stoponlevel = False
        emailadd = ""
        uname = ""
        position = ""
        SqlCmd = "Select name,signprice,position,email,signlevel from dbo.[user] where id='" & sid_create & "'"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        dr.Read()
        id = dr(4) '發送人之下一個簽單人
        InsertSignOffProcessRecord(sid_create, Session("s_name"), dr(1), dr(2), 1, dr(3), 0, 1) '發送人
        dr.Close()
        conn.Close()

        For i = 2 To signpersonmaxrow + 1
            SqlCmd = "select name,signprice,position,email,topsignoffs from dbo.[user] where id='" & id & "'"
            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
            If (dr.HasRows) Then
                dr.Read()
                uname = dr(0)
                signprice = dr(1)
                position = dr(2)
                emailadd = dr(3)
                topsignoffs = dr(4)
            End If
            dr.Close()
            conn.Close()
            If (id = "NA") Then
                Exit For
            Else
                If (topsignoffs = 0) Then
                    '寫入簽核人員表@XSPWT
                    InsertSignOffProcessRecord(id, uname, signprice, position, i, emailadd, 0, 1)
                    If (sfid > 50 And sfid < 80) Then
                        If (signprice >= (CDbl(TxtPrice.Text) * currate)) Then
                            stoponlevel = True
                            Exit For
                        End If
                    End If
                Else
                    If (i = 2) Then '如果送審人是部門主管 , 則因下一主管是global審核人會如上述跳掉,故在此加入 ,若之後在form type內訂審核人有重覆,則在那時會跳掉
                        InsertSignOffProcessRecord(id, uname, signprice, position, i, emailadd, 0, 1)
                        If (sfid > 50 And sfid < 80) Then
                            If (signprice >= (CDbl(TxtPrice.Text) * currate)) Then
                                stoponlevel = True
                                Exit For
                            End If
                        End If
                    End If
                    Exit For
                End If
            End If
            SqlCmd = "select signlevel from dbo.[user] where id='" & id & "'"
            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
            dr.Read()
            id = dr(0) '再下一個簽單人
            dr.Close()
            conn.Close()
        Next
        Return stoponlevel
    End Function
    Sub PutDataToCTForm(uid As String, prop As Integer) '123456
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim idstr As String
        'If (method = "copy") Then
        If (ResignedCheck(uid, 1) = True) Then
            CType(CT.FindControl("ddl_user_" & gseq + 1), DropDownList).Items.Clear()
            SqlCmd = "select id,name,position from dbo.[user] where email<>'' order by branch,grp"
            drL = CommUtil.SelectLocalSqlUsingDr(SqlCmd, connL) 'modify
            If (drL.HasRows) Then
                Do While (drL.Read())
                    idstr = drL(0) & " " & drL(1) & " " & drL(2)
                    CType(CT.FindControl("ddl_user_" & gseq + 1), DropDownList).Items.Add(idstr)
                Loop
            End If
            drL.Close()
            connL.Close()
            CT.Rows(gseq + 1).Cells(5).Text = 0
            CT.Rows(gseq + 1).Cells(6).Text = "已離職,請管理員修正內定表(此送審會忽略)"
        Else
            CT.Rows(gseq + 1).Cells(5).Text = 1
        End If
        'End If
        SqlCmd = "select T0.name,T0.position from dbo.[User] T0 where T0.id='" & uid & "'"
        drL = CommUtil.SelectLocalSqlUsingDr(SqlCmd, connL)
        drL.Read()
        CType(CT.FindControl("ddl_user_" & gseq + 1), DropDownList).SelectedValue = uid & " " & drL(0) & " " & drL(1)
        If (sfid <> 101) Then
            CType(CT.FindControl("ddl_user_" & gseq + 1), DropDownList).Enabled = False
        End If
        CT.Rows(gseq + 1).Cells(1).Text = drL(0)

        CT.Rows(gseq + 1).Cells(2).Text = drL(1)

        drL.Close()
        connL.Close()

        CType(CT.FindControl("rbl_prop_" & gseq + 1), RadioButtonList).SelectedValue = prop
        If (sfid <> 101) Then
            CType(CT.FindControl("rbl_prop_" & gseq + 1), RadioButtonList).Enabled = False
        End If
        If (prop = 0) Then
            CType(CT.FindControl("txt_seq_" & gseq + 1), TextBox).Text = gseq
        Else
            CType(CT.FindControl("txt_seq_" & gseq + 1), TextBox).Text = 0
        End If
        If (sfid <> 101) Then
            CType(CT.FindControl("txt_seq_" & gseq + 1), TextBox).Enabled = False
        End If
    End Sub
    Sub LetRestCTFormRecordToBeInformed()
        Dim i As Integer
        For i = gseq To signpersonmaxrow + 1
            'If (sftype <> 2) Then
            'CType(CT.FindControl("rbl_prop_" & i + 1), RadioButtonList).SelectedValue = 2
            CType(CT.FindControl("rbl_prop_" & i), RadioButtonList).Enabled = False

            'CType(CT.FindControl("txt_seq_" & i + 1), TextBox).Text = 0
            CType(CT.FindControl("txt_seq_" & i), TextBox).Enabled = False
            'End If
        Next
    End Sub
    Function GenSinoffProcessByLevelToCT()
        '此為內部門流程所經人員
        '用signlevel 一直找下去 , 直到碰到id是外部審核人員或核簽金額(有金額之表單)為止
        '之後會根據每種表單所設置之流程人員 (@XSPMT 簽核人員主檔)或核簽金額(有金額之表單), 串接成完整流程
        Dim id, uname, position, emailadd As String
        Dim i As Integer
        Dim topsignoffs As Integer
        Dim signprice As Long
        Dim stoponlevel As Boolean
        Dim currate As Single
        i = 1
        currate = 1
        If (DDLDollorUnit.SelectedIndex <> 0) Then
            If (DDLDollorUnit.SelectedValue <> "NTD") Then
                SqlCmd = "SELECT T0.[Rate], T0.[RateDate] FROM ORTT T0 where T0.[Currency]='" & DDLDollorUnit.SelectedValue & "' order by ratedate desc"
                dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                dr.Read()
                currate = dr(0)
                dr.Close()
                conn.Close()
            End If
        End If
        stoponlevel = False
        emailadd = ""
        uname = ""
        position = ""

        SqlCmd = "Select name,signprice,position,email,signlevel from dbo.[user] where id='" & sid_create & "'"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        dr.Read()
        id = dr(4) '發送人之下一個簽單人
        'InsertSignOffProcessRecord(sid_create, Session("s_name"), dr(1), dr(2), 1, dr(3), 0, 1) '發送人
        'PutDataToCTForm(sid_create, 0) '發送人===> sender 不show 在CT form , 這樣sftype=1,3 後續Sub 才能共用
        'gseq = gseq + 1
        dr.Close()
        conn.Close()

        For i = 2 To signpersonmaxrow + 1
            SqlCmd = "select name,signprice,position,email,topsignoffs from dbo.[user] where id='" & id & "'"
            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
            If (dr.HasRows) Then
                dr.Read()
                uname = dr(0)
                signprice = dr(1)
                position = dr(2)
                emailadd = dr(3)
                topsignoffs = dr(4)
            End If
            dr.Close()
            conn.Close()
            If (id = "NA") Then
                Exit For
            Else
                If (topsignoffs = 0) Then
                    '寫入簽核人員表@XSPWT
                    'InsertSignOffProcessRecord(id, uname, signprice, position, i, emailadd, 0, 1)
                    PutDataToCTForm(id, 0)
                    gseq = gseq + 1
                    If (sfid > 50 And sfid < 80) Then
                        If (signprice >= (CDbl(TxtPrice.Text) * currate)) Then
                            stoponlevel = True
                            Exit For
                        End If
                    End If
                Else
                    If (i = 2) Then '如果送審人是部門主管 , 則因下一主管是global審核人會如上述跳掉,故在此加入 ,若之後在form type內訂審核人有重覆,則在那時會跳掉
                        'InsertSignOffProcessRecord(id, uname, signprice, position, i, emailadd, 0, 1)
                        PutDataToCTForm(id, 0)
                        gseq = gseq + 1
                        If (sfid > 50 And sfid < 80) Then
                            If (signprice >= (CDbl(TxtPrice.Text) * currate)) Then
                                stoponlevel = True
                                Exit For
                            End If
                        End If
                    End If
                    Exit For
                End If
            End If
            SqlCmd = "select signlevel from dbo.[user] where id='" & id & "'"
            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
            dr.Read()
            id = dr(0) '再下一個簽單人
            dr.Close()
            conn.Close()
        Next
        Return stoponlevel
    End Function
    Sub GenSinoffProcessByFormType(stoponlevel As Boolean)
        Dim seq As Integer
        Dim signprice As Double
        'Dim leveltypeid As String
        Dim currate As Single
        Dim skipflow As Boolean
        Dim connLocal As New SqlConnection
        If (stoponlevel = False) Then
            currate = 1
            If (DDLDollorUnit.SelectedValue = "USD") Then
                SqlCmd = "SELECT T0.[Rate], T0.[RateDate] FROM ORTT T0 where T0.[Currency]='USD' order by ratedate desc"
                dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                dr.Read()
                currate = dr(0)
                dr.Close()
                conn.Close()
            End If
            signprice = 0
            seq = 1
            SqlCmd = "select IsNull(max(seq)+1,1) from [dbo].[@XSPWT] where docentry=" & docnum
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                dr.Read()
                seq = dr(0)
            End If
            dr.Close()
            connsap.Close()
            SqlCmd = "select uid from [dbo].[@XSPMT] T0 where T0.process=0 and T0.prop = 0 and T0.sfid=" & sfid & " order by T0.seq" '@XSPMT 簽核人員主檔
            drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap1)
            If (drsap.HasRows) Then
                Do While (drsap.Read())
                    skipflow = False
                    '判斷此簽核人員是否有部門排除
                    SqlCmd = "select count(*) from [dbo].[@XSDET] where uid='" & drsap(0) & "' and deptcode='" & Session("grp") & "' and sfid=" & sfid
                    dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                    If (dr.HasRows) Then
                        dr.Read()
                        If (dr(0) <> 0) Then
                            skipflow = True
                        End If
                    End If
                    dr.Close()
                    connsap.Close()
                    'Dim delseq As Integer
                    If (skipflow = False) Then
                        SqlCmd = "select seq from [dbo].[@XSPWT] where docentry=" & docnum & " and uid='" & drsap(0) & "'" '檢查在ByLevel產生簽核表中是否已列入FormType之簽核人,是就skip
                        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                        If (dr.HasRows) Then '發現重覆 ==>把已存在刪除 , 並把比此seq還大的都減一 ,之後會在重新以新的 seq 加入
                            dr.Read()        '==>發現重覆是可直接skip , 但因在FormType之簽核人位階較高 , 故有必要簽核順序要依後面的,故有下述處理
                            'delseq = dr(0)
                            'If (delseq <> seq - 1) Then '如果要加入之上層簽核人與之前ByLevel最後一個是不同人 , 則刪除Bylevel之人 , 若相同則不與處理(因seq會是一樣的)
                            '    SqlCmd = "delete from [dbo].[@XSPWT] where docentry=" & docnum & " and seq=" & delseq
                            '    CommUtil.SqlSapExecute("del", SqlCmd, connLocal)
                            '    connLocal.Close()
                            '    SqlCmd = "update [dbo].[@XSPWT] set seq=seq-1 where docentry=" & docnum & " and seq >" & dr(0)
                            '    CommUtil.SqlSapExecute("upd", SqlCmd, connLocal)
                            '    connLocal.Close()
                            '    If (delseq = 2) Then '如果被刪除的是 seq=2 , 則更新後之seq=2 之status要設為1
                            '        SqlCmd = "update [dbo].[@XSPWT] set status=1 where docentry=" & docnum & " and seq =2"
                            '        CommUtil.SqlSapExecute("upd", SqlCmd, connLocal)
                            '        connLocal.Close()
                            '    End If
                            '    seq = seq - 1
                            'Else
                            '    skipflow = True
                            'End If
                            '上述判斷太麻煩, 且若上層簽核與發送人同一人 , 則上述code無處理,故改為如下 , 上層與bylevel重複,則上層skip 2024/10/04
                            skipflow = True
                        End If
                        dr.Close()
                        connsap.Close()
                    End If
                    Dim topsignoffs As Integer
                    topsignoffs = 0
                    If (skipflow = False) Then
                        SqlCmd = "select name,signprice,position,email,topsignoffs from dbo.[user] where id='" & drsap(0) & "'"
                        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                        If (dr.HasRows) Then
                            dr.Read()
                            signprice = dr(1)
                            topsignoffs = dr(4)
                            '寫入簽核人員表@XSPWT
                            InsertSignOffProcessRecord(drsap(0), dr(0), dr(1), dr(2), seq, dr(3), 0, 0)
                        End If
                        dr.Close()
                        conn.Close()
                        seq = seq + 1
                        If (sfid > 51 And sfid < 80 Or (sfid = 51 And topsignoffs = 1)) Then 'sfid=51 (領料單 , 產生簽核list時,在設定表列之人員 ,如果不是全域性簽核者,不做金額判斷) 
                            If (signprice >= (CDbl(TxtPrice.Text) * currate)) Then
                                Exit Do
                            End If
                        End If
                    End If ' end of skipflow
                Loop
            End If
            drsap.Close()
            connsap1.Close()
        End If
    End Sub
    Sub GenSinoffProcessByFormTypeToCT(stoponlevel As Boolean)
        Dim seq As Integer
        Dim signprice As Double
        Dim currate As Single
        Dim skipflow As Boolean
        Dim connLocal As New SqlConnection
        Dim i, judgecount As Integer
        Dim str() As String
        i = 2
        judgecount = gseq - 1
        If (stoponlevel = False) Then
            currate = 1
            If (DDLDollorUnit.SelectedValue = "USD") Then
                SqlCmd = "SELECT T0.[Rate], T0.[RateDate] FROM ORTT T0 where T0.[Currency]='USD' order by ratedate desc"
                dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                dr.Read()
                currate = dr(0)
                dr.Close()
                conn.Close()
            End If
            signprice = 0
            seq = 1
            SqlCmd = "select IsNull(max(seq)+1,1) from [dbo].[@XSPWT] where docentry=" & docnum
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                dr.Read()
                seq = dr(0)
            End If
            dr.Close()
            connsap.Close()
            SqlCmd = "select uid from [dbo].[@XSPMT] T0 where T0.process=0 and T0.prop = 0 and T0.sfid=" & sfid & " order by T0.seq" '@XSPMT 簽核人員主檔
            drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap1)
            If (drsap.HasRows) Then
                Do While (drsap.Read())
                    skipflow = False
                    '判斷此簽核人員是否有部門排除
                    SqlCmd = "select count(*) from [dbo].[@XSDET] where uid='" & drsap(0) & "' and deptcode='" & Session("grp") & "' and sfid=" & sfid
                    dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                    If (dr.HasRows) Then
                        dr.Read()
                        If (dr(0) <> 0) Then
                            skipflow = True
                        End If
                    End If
                    dr.Close()
                    connsap.Close()
                    'Dim delseq As Integer
                    If (skipflow = False) Then
                        'SqlCmd = "select seq from [dbo].[@XSPWT] where docentry=" & docnum & " and uid='" & drsap(0) & "'" '檢查在ByLevel產生簽核表中是否已列入FormType之簽核人,是就skip
                        'dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                        'If (dr.HasRows) Then '發現重覆 ==>把已存在刪除 , 並把比此seq還大的都減一 ,之後會在重新以新的 seq 加入
                        '    dr.Read()        '==>發現重覆是可直接skip , 但因在FormType之簽核人位階較高 , 故有必要簽核順序要依後面的,故有下述處理
                        'delseq = dr(0)
                        'If (delseq <> seq - 1) Then '如果要加入之上層簽核人與之前ByLevel最後一個是不同人 , 則刪除Bylevel之人 , 若相同則不與處理(因seq會是一樣的)
                        '    SqlCmd = "delete from [dbo].[@XSPWT] where docentry=" & docnum & " and seq=" & delseq
                        '    CommUtil.SqlSapExecute("del", SqlCmd, connLocal)
                        '    connLocal.Close()
                        '    SqlCmd = "update [dbo].[@XSPWT] set seq=seq-1 where docentry=" & docnum & " and seq >" & dr(0)
                        '    CommUtil.SqlSapExecute("upd", SqlCmd, connLocal)
                        '    connLocal.Close()
                        '    If (delseq = 2) Then '如果被刪除的是 seq=2 , 則更新後之seq=2 之status要設為1
                        '        SqlCmd = "update [dbo].[@XSPWT] set status=1 where docentry=" & docnum & " and seq =2"
                        '        CommUtil.SqlSapExecute("upd", SqlCmd, connLocal)
                        '        connLocal.Close()
                        '    End If
                        '    seq = seq - 1
                        'Else
                        '    skipflow = True
                        'End If
                        '上述判斷太麻煩, 且若上層簽核與發送人同一人 , 則上述code無處理,故改為如下 , 上層與bylevel重複,則上層skip 2024/10/04
                        'skipflow = True
                        'End If
                        'dr.Close()
                        'connsap.Close()
                        For i = 2 To judgecount + 1
                            str = Split(CType(CT.FindControl("ddl_user_" & i), DropDownList).SelectedValue, " ")
                            If (str(0) = drsap(0) Or drsap(0) = sid_create) Then
                                skipflow = True
                                Exit For
                            End If
                        Next
                    End If
                    Dim topsignoffs As Integer
                    topsignoffs = 0
                    If (skipflow = False) Then
                        SqlCmd = "select name,signprice,position,email,topsignoffs from dbo.[user] where id='" & drsap(0) & "'"
                        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                        If (dr.HasRows) Then
                            dr.Read()
                            signprice = dr(1)
                            topsignoffs = dr(4)
                            '寫入簽核人員表@XSPWT
                            'InsertSignOffProcessRecord(drsap(0), dr(0), dr(1), dr(2), seq, dr(3), 0, 0)
                            PutDataToCTForm(drsap(0), 0)
                            gseq = gseq + 1
                        End If
                        dr.Close()
                        conn.Close()
                        seq = seq + 1
                        If (sfid > 51 And sfid < 80 Or (sfid = 51 And topsignoffs = 1)) Then 'sfid=51 (領料單 , 產生簽核list時,在設定表列之人員 ,如果不是全域性簽核者,不做金額判斷) 
                            If (signprice >= (CDbl(TxtPrice.Text) * currate)) Then
                                Exit Do
                            End If
                        End If
                    End If ' end of skipflow
                Loop
            End If
            drsap.Close()
            connsap1.Close()
        End If
    End Sub
    Function ArchivePersonCheck(sftype As Integer)
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim owncount As Integer
        Dim deptcode, ownid As String
        deptcode = ""
        ownid = ""
        owncount = 0
        SqlCmd = "select grp from dbo.[user] where id='" & sid_create & "'"
        drL = CommUtil.SelectLocalSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            drL.Read()
            deptcode = drL(0)
        End If
        drL.Close()
        connL.Close()
        If (sftype <> 3) Then
            SqlCmd = "select uid from [dbo].[@XSPMT] T0 where T0.prop = 1 and T0.sfid=" & sfid '@XSPMT 簽核人員主檔
        Else
            SqlCmd = "select uid from [dbo].[@XSPAT] T0 where T0.sgtype=1 and T0.prop = 1 and T0.ownid='" & sid_create & "'"
        End If
        'SqlCmd = "select uid from [dbo].[@XSPAT] T0 where T0.sgtype=1 and T0.prop = 1 and T0.ownid='" & sid_create & "'"
        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (drsap.HasRows) Then '有設歸檔人員
            If (sftype = 1) Then
                Do While (drsap.Read())
                    SqlCmd = "select count(*) from [dbo].[@XSDET] where uid='" & drsap(0) & "' and deptcode='" & deptcode & "' and sfid=" & sfid
                    drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                    If (drL.HasRows) Then  '選合適的歸檔者 (在各歸檔者之排除部門中設定,如請(採)購單屬捷豐之歸檔者, 在排除部門要把其他部門全打勾,自己部門不勾)
                        drL.Read()
                        If (drL(0) = 0) Then
                            owncount = owncount + 1
                            ownid = drsap(0)
                        End If
                    End If
                    drL.Close()
                    connL.Close()
                Loop
                If (owncount = 0) Then
                    drsap.Close()
                    connsap.Close()
                    CommUtil.ShowMsg(Me, "無歸檔者符合規則, 需有一個,請check")
                    Return ownid
                    Exit Function
                ElseIf (owncount > 1) Then
                    drsap.Close()
                    connsap.Close()
                    CommUtil.ShowMsg(Me, "有" & owncount & "個歸檔者符合規則, 只能一個,請check")
                    ownid = ""
                    Return ownid
                    Exit Function
                End If
            Else
                drsap.Read()
                ownid = drsap(0)
            End If
        Else '沒設定歸檔人員 , 故寫入發起人當歸檔人員
            ownid = Session("s_id")
        End If
        drsap.Close()
        connsap.Close()
        Return ownid
    End Function
    Sub GenSinoffProcessRecordsOfOthers(ownid As String)
        Dim seq As Integer
        Dim skip As Boolean
        SqlCmd = "select IsNull(max(seq)+1,1) from [dbo].[@XSPWT] where docentry=" & docnum
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            dr.Read()
            seq = dr(0)
        End If
        dr.Close()
        connsap.Close()
        '以下寫入歸檔人員
        SqlCmd = "select name,signprice,position,email from dbo.[user] where id='" & ownid & "'"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        If (dr.HasRows) Then
            dr.Read()
            '寫入簽核人員表@XSPWT
            InsertSignOffProcessRecord(ownid, dr(0), dr(1), dr(2), seq, dr(3), 1, 0)
        End If
        dr.Close()
        conn.Close()
        seq = seq + 1
        '以下寫入cc人員
        SqlCmd = "select uid from [dbo].[@XSPAT] T0 where T0.sgtype=1 and T0.prop = 2 and T0.ownid='" & sid_create & "'"
        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (drsap.HasRows) Then '有設cc人員
            Do While (drsap.Read())
                SqlCmd = "Select count(*) from [dbo].[@XSPWT] where uid='" & drsap(0) & "' and docentry=" & docnum
                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap1)
                dr.Read()
                skip = False
                If (dr(0) <> 0) Then
                    skip = True
                End If
                dr.Close()
                connsap1.Close()
                If (skip = False) Then
                    SqlCmd = "select name,signprice,position,email from dbo.[user] where id='" & drsap(0) & "'"
                    dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                    If (dr.HasRows) Then
                        dr.Read()
                        '寫入簽核人員表@XSPWT
                        InsertSignOffProcessRecord(drsap(0), dr(0), dr(1), dr(2), seq, dr(3), 2, 0)
                    End If
                    dr.Close()
                    conn.Close()
                    seq = seq + 1
                End If
            Loop
        End If
        drsap.Close()
        connsap.Close()
    End Sub
    'Sub GenSinoffProcessOfOthers_ORG(sftype As Integer, ownid As String)
    '    Dim seq As Integer
    '    Dim skip As Boolean
    '    SqlCmd = "select IsNull(max(seq)+1,1) from [dbo].[@XSPWT] where docentry=" & docnum
    '    dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
    '    If (dr.HasRows) Then
    '        dr.Read()
    '        seq = dr(0)
    '    End If
    '    dr.Close()
    '    connsap.Close()
    '    '以下寫入歸檔人員
    '    SqlCmd = "select name,signprice,position,email from dbo.[user] where id='" & ownid & "'"
    '    dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
    '    If (dr.HasRows) Then
    '        dr.Read()
    '        '寫入簽核人員表@XSPWT
    '        InsertSignOffProcessRecord(ownid, dr(0), dr(1), dr(2), seq, dr(3), 1, 0)
    '    End If
    '    dr.Close()
    '    conn.Close()
    '    seq = seq + 1
    '    '以下寫入cc人員
    '    If (sftype <> 3) Then
    '        SqlCmd = "select uid from [dbo].[@XSPMT] T0 where T0.prop = 2 and T0.sfid=" & sfid '@XSPMT 簽核人員主檔
    '    Else
    '        SqlCmd = "select uid from [dbo].[@XSPAT] T0 where T0.sgtype=1 and T0.prop = 2 and T0.ownid='" & sid_create & "'"
    '    End If
    '    drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
    '    If (drsap.HasRows) Then '有設cc人員
    '        Do While (drsap.Read())
    '            SqlCmd = "Select count(*) from [dbo].[@XSPWT] where uid='" & drsap(0) & "' and docentry=" & docnum
    '            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap1)
    '            dr.Read()
    '            skip = False
    '            If (dr(0) <> 0) Then
    '                skip = True
    '            End If
    '            dr.Close()
    '            connsap1.Close()
    '            If (skip = False) Then
    '                SqlCmd = "select name,signprice,position,email from dbo.[user] where id='" & drsap(0) & "'"
    '                dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
    '                If (dr.HasRows) Then
    '                    dr.Read()
    '                    '寫入簽核人員表@XSPWT
    '                    InsertSignOffProcessRecord(drsap(0), dr(0), dr(1), dr(2), seq, dr(3), 2, 0)
    '                End If
    '                dr.Close()
    '                conn.Close()
    '                seq = seq + 1
    '            End If
    '        Loop
    '    End If
    '    drsap.Close()
    '    connsap.Close()
    'End Sub
    Sub GenSinoffProcessOfOthersToCT(sftype As Integer, ownid As String)
        Dim skip As Boolean
        Dim i, judgecount As Integer
        Dim str() As String
        i = 2
        judgecount = gseq - 1

        SqlCmd = "select name,signprice,position,email from dbo.[user] where id='" & ownid & "'"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        If (dr.HasRows) Then
            dr.Read()
            PutDataToCTForm(ownid, 1)
            gseq = gseq + 1
        End If
        dr.Close()
        conn.Close()
        '以下寫入cc人員
        SqlCmd = "select uid from [dbo].[@XSPMT] T0 where T0.prop = 2 and T0.sfid=" & sfid '@XSPMT 簽核人員主檔
        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (drsap.HasRows) Then '有設cc人員
            Do While (drsap.Read())
                skip = False
                For i = 2 To judgecount + 1
                    str = Split(CType(CT.FindControl("ddl_user_" & i), DropDownList).SelectedValue, " ")
                    If (str(0) = drsap(0) Or drsap(0) = sid_create) Then
                        skip = True
                        Exit For
                    End If
                Next
                If (skip = False) Then
                    SqlCmd = "select name,signprice,position,email from dbo.[user] where id='" & drsap(0) & "'"
                    dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                    If (dr.HasRows) Then
                        dr.Read()
                        PutDataToCTForm(drsap(0), 2)
                        gseq = gseq + 1
                    End If
                    dr.Close()
                    conn.Close()
                End If
            Loop
        End If
        drsap.Close()
        connsap.Close()
    End Sub
    '以下為單獨依指定流程人員
    Sub GenSinoffProcessRecordsOfMain()
        Dim seq As Integer
        seq = 1
        SqlCmd = "Select name,signprice,position,email,signlevel from dbo.[user] where id='" & sid_create & "'"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        dr.Read()
        InsertSignOffProcessRecord(sid_create, Session("s_name"), dr(1), dr(2), seq, dr(3), 0, 0) '發送人
        dr.Close()
        conn.Close()
        seq = seq + 1
        SqlCmd = "select uid from [dbo].[@XSPAT] T0 where T0.process=0 and T0.sgtype=1 and T0.prop = 0 and T0.ownid='" & sid_create & "' order by T0.seq" '@XSPAT 簽核人員指定暫時主檔
        'SqlCmd = "select uid from [dbo].[@XSPAT] T0 where T0.sgtype=1 and T0.prop = 0 and T0.ownid='" & sid_create & "' order by T0.seq" '@XSPAT 簽核人員指定暫時主檔
        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap1)
        If (drsap.HasRows) Then
            Do While (drsap.Read()) '指定簽核人員之Table
                SqlCmd = "select name,signprice,position,email from dbo.[user] where id='" & drsap(0) & "'"
                dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                If (dr.HasRows) Then
                    dr.Read()
                    '寫入簽核人員表@XSPWT
                    InsertSignOffProcessRecord(drsap(0), dr(0), dr(1), dr(2), seq, dr(3), 0, 0)
                End If
                dr.Close()
                conn.Close()
                seq = seq + 1
            Loop
        End If
        drsap.Close()
        connsap1.Close()
    End Sub
    'Sub GenSinoffProcessByAssignment_ORG()
    '    Dim seq As Integer
    '    seq = 1
    '    SqlCmd = "Select name,signprice,position,email,signlevel from dbo.[user] where id='" & sid_create & "'"
    '    dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
    '    dr.Read()
    '    InsertSignOffProcessRecord(sid_create, Session("s_name"), dr(1), dr(2), seq, dr(3), 0, 0) '發送人
    '    dr.Close()
    '    conn.Close()
    '    seq = seq + 1
    '    SqlCmd = "select uid from [dbo].[@XSPAT] T0 where T0.process=0 and T0.sgtype=1 and T0.prop = 0 and T0.ownid='" & sid_create & "' order by T0.seq" '@XSPAT 簽核人員指定暫時主檔
    '    drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap1)
    '    If (drsap.HasRows) Then
    '        Do While (drsap.Read()) '指定簽核人員之Table
    '            SqlCmd = "select name,signprice,position,email from dbo.[user] where id='" & drsap(0) & "'"
    '            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
    '            If (dr.HasRows) Then
    '                dr.Read()
    '                '寫入簽核人員表@XSPWT
    '                InsertSignOffProcessRecord(drsap(0), dr(0), dr(1), dr(2), seq, dr(3), 0, 0)
    '            End If
    '            dr.Close()
    '            conn.Close()
    '            seq = seq + 1
    '        Loop
    '    End If
    '    drsap.Close()
    '    connsap1.Close()
    'End Sub
    Sub CopySignPersonFromMainDoc(docnum As Long)
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim uid, uname, upos, emailadd, orgsenderid, senderupos, senderemail As String
        Dim signprice, sendersignprice As Long
        Dim seq, signprop, innerloop As Integer
        orgsenderid = ""
        SqlCmd = "Select name,signprice,position,email,signlevel from dbo.[user] where id='" & sid_create & "'"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        dr.Read()
        sendersignprice = dr(1)
        senderupos = dr(2)
        senderemail = dr(3)
        InsertSignOffProcessRecord(sid_create, Session("s_name"), sendersignprice, senderupos, 1, senderemail, 0, 1) '發送人
        dr.Close()
        conn.Close()
        seq = 2
        SqlCmd = "select uid,uname,signprice,upos,seq,emailadd,signprop,innerloop from [dbo].[@XSPWT] " &
            "where docentry=" & docnum & " order by seq"
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            Do While (drL.Read())
                If (drL(4) <> 1) Then
                    If (drL(6) <> 1 Or (drL(6) = 1 And drL(0) <> orgsenderid)) Then '原簽核list 有設歸檔人
                        uid = drL(0)
                        uname = drL(1)
                        signprice = drL(2)
                        upos = drL(3)
                        'seq = drL(4)
                        emailadd = drL(5)
                        signprop = drL(6)
                        innerloop = drL(7)
                    Else
                        uid = sid_create
                        uname = Session("s_name")
                        signprice = sendersignprice
                        upos = senderupos
                        'seq = drL(4)
                        emailadd = senderemail
                        signprop = 1
                        innerloop = 0
                    End If
                    If (uid <> sid_create Or signprop = 1) Then
                        InsertSignOffProcessRecord(uid, uname, signprice, upos, seq, emailadd, signprop, innerloop)
                        seq = seq + 1
                    End If
                Else
                    orgsenderid = drL(0)
                    If (orgsenderid <> sid_create) Then
                        uid = drL(0)
                        uname = drL(1)
                        signprice = drL(2)
                        upos = drL(3)
                        emailadd = drL(5)
                        signprop = drL(6)
                        innerloop = drL(7)
                        InsertSignOffProcessRecord(uid, uname, signprice, upos, seq, emailadd, signprop, innerloop)
                        seq = seq + 1
                    End If
                End If
            Loop
        End If
        drL.Close()
        connL.Close()
    End Sub
    Sub CopySignPersonFromMainDocToCT(docnum As Long)
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim uid, orgsenderid As String
        Dim signprop As Integer
        orgsenderid = ""
        SqlCmd = "select uid,uname,signprice,upos,seq,emailadd,signprop,innerloop from [dbo].[@XSPWT] " &
            "where docentry=" & docnum & " order by seq"
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            Do While (drL.Read())
                If (drL(4) <> 1) Then
                    If (drL(6) <> 1 Or (drL(6) = 1 And drL(0) <> orgsenderid)) Then '原簽核list 有設歸檔人
                        uid = drL(0)
                        signprop = drL(6)
                    Else
                        uid = sid_create
                        signprop = 1
                    End If
                    If (uid <> sid_create Or signprop = 1) Then
                        PutDataToCTForm(uid, signprop)
                        gseq = gseq + 1
                    End If
                Else
                    orgsenderid = drL(0)
                    If (orgsenderid <> sid_create) Then
                        uid = drL(0)
                        signprop = drL(6)
                        PutDataToCTForm(uid, signprop)
                        gseq = gseq + 1
                    End If
                End If
            Loop
        End If
        drL.Close()
        connL.Close()
    End Sub
    Sub InsertSignOffProcessRecord(id As String, uname As String, signprice As Long, position As String, seq As Integer, emailadd As String, signprop As Integer, innerloop As Integer)
        Dim connL As New SqlConnection
        Dim status As Integer
        Dim signdate As String
        signdate = Format(Now(), "yyyy/MM/dd HH:mm:ss")
        'status 0:備核  1:關卡待核 2:已核  3:駁回  4: 抽回 5:已重送 6:取消  100:送審
        If (seq = 2) Then
            status = 1
        ElseIf (seq = 1) Then
            status = 100
        Else
            status = 0
        End If
        If (seq = 1) Then
            SqlCmd = "insert into [dbo].[@XSPWT] (docentry,uid,uname,upos,seq,signprice,status,signdate,emailadd,signprop,innerloop) " &
        "values(" & docnum & ",'" & id & "','" & uname & "','" & position & "'," & seq & "," & signprice & "," & status & ",'" &
                signdate & "','" & emailadd & "'," & signprop & "," & innerloop & ")"
        Else
            SqlCmd = "insert into [dbo].[@XSPWT] (docentry,uid,uname,upos,seq,signprice,status,emailadd,signprop,innerloop) " &
        "values(" & docnum & ",'" & id & "','" & uname & "','" & position & "'," & seq & "," & signprice & "," & status & ",'" & emailadd & "'," & signprop & "," & innerloop & ")"
        End If
        CommUtil.SqlSapExecute("ins", SqlCmd, connL)
        connL.Close()
    End Sub
    Protected Sub BtnRecall_Click(sender As Object, e As EventArgs)
        Dim seq As Integer
        Dim comment As String
        Dim receivedate As String
        receivedate = Format(Now(), "yyyy/MM/dd HH:mm:ss")
        comment = TxtComm.Text
        '如果不是送審人 :把@XSPWT (簽核人員Table) 上層主管status設為0 , 抽回之人status設為1 ,comment,signdate,uimagesign清掉,@XASCH 不動
        '如果是送審人 :把@XSPWT (簽核人員Table) 上層主管status設為0 ,@XASCH 之status設為'R'
        '在@XSPHT(簽核歷史Table)記錄此抽回資料
        SqlCmd = "Select seq,comment from [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum & " and uid='" & Session("s_id") & "'"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            dr.Read()
            seq = dr(0)
            'comment = dr(1)
            dr.Close()
            connsap.Close()
            '上層主管簽核狀態改為備核 'receivedate 不清除 , 留作比對
            SqlCmd = "update [dbo].[@XSPWT] set status=0 where signprop=0 and docentry=" & docnum & " and seq=" & seq + 1
            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
            connsap.Close()
            '本人簽核狀態改為關卡待核 , 其他資料清掉
            SqlCmd = "update [dbo].[@XSPWT] set status=1,receivedate='" & receivedate & "',comment='',signdate=null " &
            "where signprop=0 and docentry=" & docnum & " And seq=" & seq
            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
            connsap.Close()
            If (seq = 1) Then '送審人抽回
                SqlCmd = "update [dbo].[@XASCH] set status='R' where docnum=" & docnum
                CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                connsap.Close()
                docstatus = "R"
            End If
        End If
        '以下寫入簽核歷史資料表@XSPHT
        RecordSignFlowHistoty(comment, "抽回")
        Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&status=" & docstatus & "&docnum=" & docnum &
                          "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & sfid & "&agnid=" & agnidG &
                          "&signflowmode=" & signflowmode)
    End Sub
    Protected Sub BtnReject_Click(sender As Object, e As EventArgs)
        Dim comment As String
        Dim receivedate As String
        Dim beapproved, signfinish As Boolean
        Dim innerloop, maxseq, nowseq As Integer
        innerloop = 0
        beapproved = False
        '先check 是否此單已被覆核過
        SqlCmd = "Select status,innerloop,seq " &
                 "FROM dbo.[@XSPWT] where uid='" & Session("s_id") & "' and docentry=" & docnum & " and signprop=0"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            dr.Read()
            innerloop = dr(1)
            nowseq = dr(2)
            If (dr(0) <> 1) Then
                beapproved = True
                CommUtil.ShowMsg(Me, "此單已被他人覆核過,不予處理,進入下一單")
            End If
        Else
            CommUtil.ShowMsg(Me, Session("s_id") & " " & ds.Tables(0).Rows(Session("startindex"))("docnum") & "簽核列表中在資料庫中找不到")
        End If
        dr.Close()
        connsap.Close()
        '''''''''''''''''end
        If (beapproved = False) Then
            comment = TxtComm.Text
            If (comment = "") Then
                CommUtil.ShowMsg(Me, "!!!請在意見欄內填寫駁回或反對之原因!!!")
                Exit Sub
            End If
            SqlCmd = "Select max(seq) from [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            maxseq = dr(0)
            dr.Close()
            connsap.Close()
            receivedate = Format(Now(), "yyyy/MM/dd HH:mm:ss")
            If (innerloop = 1 Or maxseq = nowseq Or ChkReturn.Checked) Then '退回送審者
                '先把所有簽核人員status設為0 'receivedate 不清除 , 留作比對
                SqlCmd = "update [dbo].[@XSPWT] set status=0,comment='',signdate=null " &
                        "where signprop=0 and docentry=" & docnum
                CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                connsap.Close()
                '送審人status=1(待核)
                SqlCmd = "update [dbo].[@XSPWT] set status=1,receivedate='" & receivedate & "' " &
                        "where signprop=0 and docentry=" & docnum & " and seq=1"
                CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                connsap.Close()
                '將簽退人員之status改為3 ==>= 是否需要 , 如需要在此員簽核欄會有X圖案 , 直到其再次簽核時
                SqlCmd = "update [dbo].[@XSPWT] set status=3 " &
                        "where signprop=0 and docentry=" & docnum & " and uid='" & Session("s_id") & "'"
                CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                connsap.Close()
                SqlCmd = "update [dbo].[@XASCH] set status='B' where docnum=" & docnum
                CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                connsap.Close()
                docstatus = "B"
                '在@XSPHT(簽核歷史Table)記錄此退回資料
                RecordSignFlowHistoty(comment, "駁回")
                Email_SignOffFlow("reject", 0)
            Else '表示反對 , 但還是進入下關
                '上層主管簽核狀態改為關卡待核
                SqlCmd = "update [dbo].[@XSPWT] set status=1,receivedate='" & receivedate & "' where signprop=0 and docentry=" & docnum & " and seq=" & nowseq + 1
                CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                connsap.Close()
                '本身status 設為3 ,其他欄位存檔
                SqlCmd = "update [dbo].[@XSPWT] set status=3,comment='" & comment & "',signdate='" & receivedate & "',agnid='" & agnidG & "' " &
                     "where signprop=0 and docentry=" & docnum & " And seq=" & nowseq
                CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                connsap.Close()
                RecordSignFlowHistoty(comment, "反對")
                Email_SignOffFlow("against", 0)
            End If
        End If

        If (actmode = "signoff" Or actmode = "recycle" Or actmode = "recycle_login" Or actmode = "signoff_login") Then
            ds.Tables(0).Rows(Session("startindex"))("signoffflag") = 1
            Session("ds") = ds
            signfinish = FindNextSignDoc()
            If (signfinish = False) Then
                Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&actmode=recycle&status=" & ds.Tables(0).Rows(Session("startindex"))("status") &
                                "&docnum=" & ds.Tables(0).Rows(Session("startindex"))("docnum") &
                                "&formtypeindex=" & formtypeindex & "&formstatusindex=" & formstatusindex &
                                "&sfid=" & ds.Tables(0).Rows(Session("startindex"))("sfid") & "&agnid=" & agnidG & "&signflowmode=" & signflowmode)
            Else
                If (actmode = "recycle_login" Or actmode = "signoff_login") Then
                    Response.Redirect("~/signoff/signoff.aspx?smid=sg&smode=1&act=signfinish&signflowmode=" & signflowmode)
                Else
                    Response.Redirect("~/usermgm/logout.aspx?act=signfinish")
                End If
            End If
        Else
            Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&status=B&docnum=" & docnum &
              "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & sfid & "&agnid=" & agnidG & "&signflowmode=" & signflowmode)
        End If
    End Sub
    Protected Sub BtnApproval_Click(sender As Object, e As EventArgs)
        '把@XSPWT (簽核人員Table) 上層主管status設為1 ,本身status 設為2 ,其他欄位存檔
        '@XASCH 之status 若是最後一關 , status設為'F'
        Dim seq, maxseq As Integer
        Dim comment, reason As String
        Dim signdate As String
        Dim beapproved, signfinish As Boolean
        beapproved = False
        '先check 是否此單已被覆核過
        SqlCmd = "Select status " &
                 "FROM dbo.[@XSPWT] where uid='" & Session("s_id") & "' and docentry=" & docnum & " and signprop=0"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            dr.Read()
            If (dr(0) <> 1) Then
                beapproved = True
                CommUtil.ShowMsg(Me, "此單已被他人覆核過,不予處理,進入下一單")
                'Threading.Thread.Sleep(1000)
            End If
        Else
            CommUtil.ShowMsg(Me, Session("s_id") & " " & ds.Tables(0).Rows(Session("startindex"))("docnum") & "簽核列表中在資料庫中找不到")
        End If
        dr.Close()
        connsap.Close()
        '''''''''''''''''end
        If (beapproved = False) Then
            signdate = Format(Now(), "yyyy/MM/dd HH:mm:ss")
            comment = TxtComm.Text
            reason = ""
            SqlCmd = "Select max(seq) from [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            maxseq = dr(0)
            dr.Close()
            connsap.Close()

            SqlCmd = "Select seq from [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum & " and uid='" & Session("s_id") & "'"
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                dr.Read()
                seq = dr(0)
                dr.Close()
                connsap.Close()
                If (seq < maxseq) Then
                    '上層主管簽核狀態改為關卡待核
                    SqlCmd = "update [dbo].[@XSPWT] set status=1,receivedate='" & signdate & "' where signprop=0 and docentry=" & docnum & " and seq=" & seq + 1
                    CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                    connsap.Close()
                End If
                '本身status 設為2 ,其他欄位存檔
                SqlCmd = "update [dbo].[@XSPWT] set status=2,comment='" & comment & "',signdate='" & signdate & "',agnid='" & agnidG & "' " &
                     "where signprop=0 and docentry=" & docnum & " And seq=" & seq
                CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                connsap.Close()
                '@XASCH 之status 若是最後一關 , status設為'F'
                If (maxseq = seq) Then '最後一關
                    SqlCmd = "update [dbo].[@XASCH] set status='F' where docnum=" & docnum
                    CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                    connsap.Close()
                    reason = "結案"
                    docstatus = "F"
                    '設歸檔人員 status=1及 receivedate
                    SqlCmd = "update [dbo].[@XSPWT] set status=1,receivedate='" & signdate & "' where signprop=1 and docentry=" & docnum & " And seq=" & seq + 1
                    CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                    connsap.Close()
                    '設知悉人員 receivedate
                    SqlCmd = "update [dbo].[@XSPWT] set status=1,receivedate='" & signdate & "' where signprop=2 and docentry=" & docnum
                    CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                    connsap.Close()
                    '若為加簽單 , 則需update其單號到母單號
                    If (sfid = 100 Or sfid = 101) Then
                        Dim mainattadoc As String
                        SqlCmd = "Select attadoc from [dbo].[@XASCH] where docnum=" & CLng(TxtAttaDoc.Text)
                        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                        dr.Read()
                        mainattadoc = dr(0)
                        dr.Close()
                        connsap.Close()
                        If (mainattadoc = "NA") Then
                            SqlCmd = "update [dbo].[@XASCH] set attadoc='" & CStr(docnum) & "' where docnum=" & CLng(TxtAttaDoc.Text)
                            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                            connsap.Close()
                        Else
                            mainattadoc = mainattadoc & "_" & CStr(docnum)
                            SqlCmd = "update [dbo].[@XASCH] set attadoc='" & mainattadoc & "' where docnum=" & CLng(TxtAttaDoc.Text)
                            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                            connsap.Close()
                        End If
                        'If (sfid = 101) Then '已放到產生pdf 後才執行 , 因如此才能得到正確此次返還數量
                        '    '將設定之返還數量加到rtnqty
                        '    If (TxtAttaDoc.Text <> "NA" And TxtAttaDoc.Text <> "") Then
                        '        SqlCmd = "update [dbo].[@XSMLS] set rtnqty=rtnqty+nowrtnqty " &
                        '            " where docentry=" & CLng(TxtAttaDoc.Text) & " and head=0"
                        '        CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                        '        connsap.Close()
                        '    End If
                        '    '將設定之返還數量歸0
                        '    If (TxtAttaDoc.Text <> "NA" And TxtAttaDoc.Text <> "") Then
                        '        SqlCmd = "update [dbo].[@XSMLS] set nowrtnqty=0 " &
                        '            " where docentry=" & CLng(TxtAttaDoc.Text) & " and head=0"
                        '        CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                        '        connsap.Close()
                        '    End If
                        'End If
                    End If
                Else
                    reason = "核准"
                End If
            End If
            '在@XSPHT(簽核歷史Table)記錄此核准資料
            RecordSignFlowHistoty(comment, reason)
            Email_SignOffFlow("approval", 0)
        End If
        Dim status As String
        If (actmode = "signoff" Or actmode = "recycle" Or actmode = "recycle_login" Or actmode = "signoff_login") Then
            ds.Tables(0).Rows(Session("startindex"))("signoffflag") = 1 '設定該員簽核表單list中 , 此單設為已簽
            Session("ds") = ds
            signfinish = FindNextSignDoc()
            status = ds.Tables(0).Rows(Session("startindex"))("status")
            If (signfinish = False) Then
                Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&actmode=recycle&status=" & status &
                                "&docnum=" & ds.Tables(0).Rows(Session("startindex"))("docnum") &
                                "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & ds.Tables(0).Rows(Session("startindex"))("sfid") &
                                "&agnid=" & agnidG & "&signflowmode=" & signflowmode)
            Else
                If (actmode = "recycle_login" Or actmode = "signoff_login") Then
                    Response.Redirect("~/signoff/signoff.aspx?smid=sg&smode=1&act=signfinish&signflowmode=" & signflowmode)
                Else
                    Response.Redirect("~/usermgm/logout.aspx?act=signfinish")
                End If
            End If
        Else
            Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&status=" & docstatus & "&docnum=" & docnum &
              "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & sfid & "&agnid=" & agnidG & "&signflowmode=" & signflowmode)
        End If
        'ViewState("sfid") = sfid
    End Sub
    Protected Sub BtnSkip_Click(sender As Object, e As EventArgs)
        '把@XSPWT (簽核人員Table) 上層主管status設為1 ,本身status 設為3 ,其他欄位存檔
        Dim seq, maxseq As Integer
        Dim comment, reason As String
        Dim signdate As String
        Dim beapproved As Boolean
        Dim skipid As String
        beapproved = False
        skipid = Request.QueryString("skipid")
        '先check 是否此單已被覆核過
        SqlCmd = "Select status " &
                 "FROM dbo.[@XSPWT] where uid='" & skipid & "' and docentry=" & docnum & " and signprop=0"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            dr.Read()
            If (dr(0) <> 1) Then
                beapproved = True
                CommUtil.ShowMsg(Me, "此單已被他人覆核過,不予處理,進入下一單")
            End If
        Else
            CommUtil.ShowMsg(Me, skipid & " 表單號:" & docnum & "簽核列表中在資料庫中找不到")
        End If
        dr.Close()
        connsap.Close()
        '''''''''''''''''end
        If (beapproved = False) Then
            signdate = Format(Now(), "yyyy/MM/dd HH:mm:ss")
            comment = "管理者執行跳過簽核 " & TxtComm.Text
            reason = ""
            SqlCmd = "Select max(seq) from [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            maxseq = dr(0)
            dr.Close()
            connsap.Close()

            SqlCmd = "Select seq from [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum & " and uid='" & skipid & "'"
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                dr.Read()
                seq = dr(0)
                dr.Close()
                connsap.Close()
                If (seq < maxseq) Then
                    '上層主管簽核狀態改為關卡待核
                    SqlCmd = "update [dbo].[@XSPWT] set status=1,receivedate='" & signdate & "' where signprop=0 and docentry=" & docnum & " and seq=" & seq + 1
                    CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                    connsap.Close()
                End If
                '本身status 設為10 ,其他欄位存檔
                SqlCmd = "update [dbo].[@XSPWT] set status=10,signdate='" & signdate & "',comment='" & comment & "' " &
                     "where signprop=0 and docentry=" & docnum & " And seq=" & seq
                CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                connsap.Close()
            Else
                dr.Close()
                connsap.Close()
            End If
            '在@XSPHT(簽核歷史Table)記錄此核准資料
            RecordSignFlowHistoty(comment, "跳過簽核")
            Email_SignOffFlow("skip", 1)
            Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&act=skipok&status=" & docstatus & "&docnum=" & docnum &
                              "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & sfid & "&agnid=" & agnidG & "&signflowmode=" & signflowmode)
        Else
            CommUtil.ShowMsg(Me, "此單剛已被覆核過")
        End If

    End Sub
    Function FindNextSignDoc()
        Dim i, j As Integer
        Dim status As String
        Dim nonesign As Boolean
        nonesign = True
        For i = 0 To Session("sgcount") - 1
            j = Session("startindex") + 1
            If (j >= Session("sgcount")) Then
                j = 0
            End If
            If (ds.Tables(0).Rows(j)("signoffflag") = 0) Then
                nonesign = False
                Session("startindex") = j
                '以下若有代理人啟動的話 , 為防止2人同時在簽核 , 故需檢查下一簽核單是否以被覆核過
                status = ds.Tables(0).Rows(Session("startindex"))("status")
                If (status = "A" Or status = "E" Or status = "D" Or status = "R" Or status = "B") Then '如果是未送審
                    SqlCmd = "Select status " &
                             "FROM dbo.[@XASCH] where sid='" & Session("s_id") & "' and docnum=" & ds.Tables(0).Rows(Session("startindex"))("docnum")
                    dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                    If (dr.HasRows) Then
                        dr.Read()
                        If (dr(0) <> status) Then
                            ds.Tables(0).Rows(Session("startindex"))("signoffflag") = 1 '設定該員簽核表單list中 , 此單設為已簽
                            Session("ds") = ds
                        Else
                            dr.Close()
                            connsap.Close()
                            Exit For
                        End If
                    Else
                        CommUtil.ShowMsg(Me, Session("s_id") & " " & ds.Tables(0).Rows(Session("startindex"))("docnum") & "簽核列表中在資料庫中找不到")
                    End If
                    dr.Close()
                    connsap.Close()
                Else
                    SqlCmd = "Select status " &
                             "FROM dbo.[@XSPWT] where uid='" & Session("s_id") & "' and docentry=" & ds.Tables(0).Rows(Session("startindex"))("docnum") &
                             " and status<>100"
                    dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                    If (dr.HasRows) Then
                        dr.Read()
                        If (dr(0) <> 1) Then
                            ds.Tables(0).Rows(Session("startindex"))("signoffflag") = 1 '設定該員簽核表單list中 , 此單設為已簽
                            Session("ds") = ds
                        Else
                            dr.Close()
                            connsap.Close()
                            Exit For
                        End If
                    Else
                        CommUtil.ShowMsg(Me, Session("s_id") & " " & ds.Tables(0).Rows(Session("startindex"))("docnum") & "簽核列表中在資料庫中找不到")
                    End If
                    dr.Close()
                    connsap.Close()
                End If
            End If
        Next
        'If (nonesign) Then
        'CommUtil.ShowMsg(Me, "已簽核完畢")
        'End If
        Return nonesign
    End Function
    Protected Sub BtnCancel_Click(sender As Object, e As EventArgs)
        Dim comment As String
        Dim beapproved, signfinish As Boolean
        beapproved = False
        '先check 是否此單已被覆核過
        SqlCmd = "Select status " &
                 "FROM dbo.[@XASCH] where docnum=" & docnum
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            dr.Read()
            If (dr(0) <> docstatus) Then
                beapproved = True
                CommUtil.ShowMsg(Me, "此單已被他人覆核過,不予處理,進入下一單")
            End If
        Else
            CommUtil.ShowMsg(Me, Session("s_id") & " " & ds.Tables(0).Rows(Session("startindex"))("docnum") & "簽核列表中在資料庫中找不到")
        End If
        dr.Close()
        connsap.Close()
        '''''''''''''''''end
        If (beapproved = False) Then
            comment = TxtComm.Text
            '@XASCH 之status設為'R'
            '在@XSPHT(簽核歷史Table)記錄此取消資料

            SqlCmd = "update [dbo].[@XASCH] Set status='C' where docnum=" & docnum
            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
            connsap.Close()
            docstatus = "C"
            '將送審者簽核狀態改為取消
            SqlCmd = "update [dbo].[@XSPWT] set status=5 " &
                        "where docentry=" & docnum & " and seq=1"
            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
            connsap.Close()

            '以下寫入簽核歷史資料表@XSPHT
            RecordSignFlowHistoty(comment, "作廢")
            '將設定之返還數量歸0
            If (sfid = 101) Then
                If (TxtAttaDoc.Text <> "NA" And TxtAttaDoc.Text <> "") Then
                    SqlCmd = "update [dbo].[@XSMLS] set nowrtnqty=0 " &
                        " where docentry=" & CLng(TxtAttaDoc.Text) & " and head=0"
                    CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                    connsap.Close()
                End If
            End If
        End If
        Dim status As String
        If (actmode = "signoff" Or actmode = "recycle" Or actmode = "recycle_login" Or actmode = "signoff_login") Then
            ds.Tables(0).Rows(Session("startindex"))("signoffflag") = 1 '設定該員簽核表單list中 , 此單設為已簽
            Session("ds") = ds
            signfinish = FindNextSignDoc()
            status = ds.Tables(0).Rows(Session("startindex"))("status")
            If (signfinish = False) Then
                Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&actmode=recycle&status=" & status &
                                "&docnum=" & ds.Tables(0).Rows(Session("startindex"))("docnum") &
                                "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & ds.Tables(0).Rows(Session("startindex"))("sfid") &
                                "&agnid=" & agnidG & "&signflowmode=" & signflowmode)
            Else
                If (actmode = "recycle_login" Or actmode = "signoff_login") Then
                    Response.Redirect("~/signoff/signoff.aspx?smid=sg&smode=1&act=signfinish&signflowmode=" & signflowmode)
                Else
                    Response.Redirect("~/usermgm/logout.aspx?act=signfinish")
                End If
            End If
        Else
            Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&status=" & docstatus & "&docnum=" & docnum &
              "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & sfid & "&agnid=" & agnidG & "&signflowmode=" & signflowmode)
        End If
    End Sub

    Protected Sub BtnArchieve_Click(sender As Object, e As EventArgs)
        Dim comment As String
        Dim signdate As String
        Dim beapproved, signfinish, informincharge As Boolean
        beapproved = False
        informincharge = False
        '
        SqlCmd = "Select incharge " &
                 "FROM dbo.[@XSTDT] where docentry=" & docnum
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            dr.Read()
            If (dr(0) = "") Then
                CommUtil.ShowMsg(Me, "請先設定此單之追蹤單據負責人後再執行歸檔")
                dr.Close()
                connsap.Close()
                Response.Redirect("~/signoff/signofftodo.aspx?smid=sg&smode=6&actstr=informsetincharge&docentry=" & docnum & "&actmode=" & actmode)
                'Exit Sub
            Else
                informincharge = True
            End If
        End If
        dr.Close()
        connsap.Close()
        '先check 是否此單已被覆核過
        SqlCmd = "Select status " &
                 "FROM dbo.[@XSPWT] where signprop=1 and uid='" & Session("s_id") & "' and docentry=" & docnum
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            dr.Read()
            If (dr(0) <> 1) Then
                beapproved = True
                CommUtil.ShowMsg(Me, "此單已被他人覆核過,不予處理,進入下一單")
                'MsgBox("此單已被他人覆核過,不予處理,進入下一單")
            End If
        Else
            CommUtil.ShowMsg(Me, Session("s_id") & " " & ds.Tables(0).Rows(Session("startindex"))("docnum") & "簽核列表中在資料庫中找不到")
        End If
        dr.Close()
        connsap.Close()
        '''''''''''''''''end
        If (beapproved = False) Then
            signdate = Format(Now(), "yyyy/MM/dd HH:mm:ss")
            comment = TxtComm.Text
            '@XASCH 之status設為'T'
            '在@XSPHT(簽核歷史Table)記錄此歸檔資料
            '將歸檔者簽核狀態改為已核
            SqlCmd = "update [dbo].[@XSPWT] set status=103,comment='" & comment & "',signdate='" & signdate & "',agnid='" & agnidG & "' " &
                        "where docentry=" & docnum & " and signprop=1"
            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
            connsap.Close()

            SqlCmd = "update [dbo].[@XASCH] set status='T' where docnum=" & docnum
            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
            connsap.Close()
            docstatus = "T"
            '以下寫入簽核歷史資料表@XSPHT
            RecordSignFlowHistoty(comment, "歸檔")
            If (informincharge) Then
                CommSignOff.Email_InformInCharge(docnum, sfid, url)
            End If
        End If
        Dim status As String
        If (actmode = "signoff" Or actmode = "recycle" Or actmode = "recycle_login" Or actmode = "signoff_login") Then
            ds.Tables(0).Rows(Session("startindex"))("signoffflag") = 1 '設定該員簽核表單list中 , 此單設為已簽
            Session("ds") = ds
            signfinish = FindNextSignDoc()
            status = ds.Tables(0).Rows(Session("startindex"))("status")
            If (signfinish = False) Then
                Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&actmode=recycle&status=" & status & "&docnum=" & ds.Tables(0).Rows(Session("startindex"))("docnum") &
                          "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & ds.Tables(0).Rows(Session("startindex"))("sfid") &
                          "&agnid=" & agnidG & "&signflowmode=" & signflowmode)
            Else
                If (actmode = "recycle_login" Or actmode = "signoff_login") Then
                    Response.Redirect("~/signoff/signoff.aspx?smid=sg&smode=1&act=signfinish&signflowmode=" & signflowmode)
                Else
                    Response.Redirect("~/usermgm/logout.aspx?act=signfinish")
                End If
            End If
        Else
            Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&status=T&docnum=" & docnum &
              "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & sfid & "&agnid=" & agnidG & "&signflowmode=" & signflowmode)
        End If
    End Sub
    Protected Sub BtnBeInformed_Click(sender As Object, e As EventArgs)
        Dim status As String
        Dim signfinish As Boolean
        Dim signdate As String
        signdate = Format(Now(), "yyyy/MM/dd HH:mm:ss")
        '把status update為104
        SqlCmd = "update [dbo].[@XSPWT] set status=104,signdate='" & signdate & "' where signprop=2 and docentry=" & docnum & " and uid='" & Session("s_id") & "'"
            CommUtil.SqlSapExecute("upd", SqlCmd, connsap1)
            connsap1.Close()
        '寫入History
        '在@XSPHT(簽核歷史Table)記錄此核准資料
        RecordSignFlowHistoty("", "已知悉")

        If (actmode = "signoff" Or actmode = "recycle" Or actmode = "recycle_login" Or actmode = "signoff_login") Then
            ds.Tables(0).Rows(Session("startindex"))("signoffflag") = 1 '設定該員簽核表單list中 , 此單設為已簽
            Session("ds") = ds
            signfinish = FindNextSignDoc()
            status = ds.Tables(0).Rows(Session("startindex"))("status")
            If (signfinish = False) Then
                Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&actmode=recycle&status=" & status & "&docnum=" & ds.Tables(0).Rows(Session("startindex"))("docnum") &
                          "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & sfid & "&agnid=" & agnidG & "&signflowmode=" & signflowmode)
            Else
                If (actmode = "recycle_login" Or actmode = "signoff_login") Then
                    Response.Redirect("~/signoff/signoff.aspx?smid=sg&smode=1&act=signfinish&signflowmode=" & signflowmode)
                Else
                    Response.Redirect("~/usermgm/logout.aspx?act=signfinish")
                End If
            End If
        Else
            Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&status=T&docnum=" & docnum &
              "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & sfid & "&agnid=" & agnidG & "&signflowmode=" & signflowmode)
        End If
    End Sub
    Protected Sub BtnNext_Click(sender As Object, e As EventArgs)
        Session("startindex") = Session("startindex") + 1  ' begin 0 To Session("sgcount")-1
        If (Session("startindex") >= (Session("sgcount"))) Then
            Session("startindex") = Session("sgcount") - 1
        End If
        Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&actmode=recycle&status=" & ds.Tables(0).Rows(Session("startindex"))("status") &
                    "&docnum=" & ds.Tables(0).Rows(Session("startindex"))("docnum") & "&signoffflag=" & ds.Tables(0).Rows(Session("startindex"))("signoffflag") &
                    "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & ds.Tables(0).Rows(Session("startindex"))("sfid") & "&agnid=" & agnidG & "&signflowmode=" & signflowmode)
    End Sub
    Protected Sub BtnLast_Click(sender As Object, e As EventArgs)
        Session("startindex") = Session("startindex") - 1
        If (Session("startindex") < 0) Then
            Session("startindex") = 0
        End If
        ViewState("sid_create") = sid_create
        Response.Redirect("~/signoff/cLsignoff.aspx?smid=sg&smode=2&actmode=recycle&status=" & ds.Tables(0).Rows(Session("startindex"))("status") &
                    "&docnum=" & ds.Tables(0).Rows(Session("startindex"))("docnum") & "&signoffflag=" & ds.Tables(0).Rows(Session("startindex"))("signoffflag") &
                    "&formtypeindex=" & formtypeindex & "&formstatusindex=0&sfid=" & ds.Tables(0).Rows(Session("startindex"))("sfid") & "&agnid=" & agnidG & "&signflowmode=" & signflowmode)
    End Sub

    Sub RecordSignFlowHistoty(comment As String, reason As String)
        Dim flowseq As Integer
        Dim signdate As String
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim agnname As String
        Dim signid, signname As String
        signname = ""
        If (reason = "跳過簽核") Then
            signid = Request.QueryString("skipid")
        Else
            signid = Session("s_id")
        End If
        SqlCmd = "select name from dbo.[user] where id='" & signid & "'"
        drL = CommUtil.SelectLocalSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            drL.Read()
            signname = drL(0)
        End If
        drL.Close()
        connL.Close()
        signdate = Format(Now(), "yyyy/MM/dd HH:mm:ss")
        SqlCmd = "Select IsNull(Max(flowseq),0) from [dbo].[@XSPHT] where docentry=" & docnum
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        dr.Read()
        flowseq = dr(0) + 1
        dr.Close()
        connsap.Close()
        agnname = ""
        If (agnidG <> "") Then
            SqlCmd = "select name from dbo.[user] where id='" & agnidG & "'"
            drL = CommUtil.SelectLocalSqlUsingDr(SqlCmd, connL)
            If (drL.HasRows) Then
                drL.Read()
                agnname = drL(0)
            End If
            drL.Close()
            connL.Close()
        End If
        SqlCmd = "insert into [dbo].[@XSPHT] (docentry,uid,uname,flowseq,signdate,status,comment,agnname) " &
        "values(" & docnum & ",'" & signid & "','" & signname & "'," & flowseq &
        ",'" & signdate & "','" & reason & "','" & comment & "','" & agnname & "')"
        CommUtil.SqlSapExecute("ins", SqlCmd, connsap)
        connsap.Close()
    End Sub
    Sub RecordSignFlowHistotySfid100_101(comment As String, reason As String)
        Dim flowseq As Integer
        Dim signdate As String
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim agnname As String
        Dim signid, signname As String
        signid = ""
        signname = ""
        SqlCmd = "select sid,sname from dbo.[@XASCH] where docnum=" & docnum
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            drL.Read()
            signid = drL(0)
            signname = drL(1)
        End If
        drL.Close()
        connL.Close()
        signdate = Format(Now(), "yyyy/MM/dd HH:mm:ss")
        SqlCmd = "Select IsNull(Max(flowseq),0) from [dbo].[@XSPHT] where docentry=" & CLng(TxtAttaDoc.Text)
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        dr.Read()
        flowseq = dr(0) + 1
        dr.Close()
        connsap.Close()
        agnname = ""
        SqlCmd = "insert into [dbo].[@XSPHT] (docentry,uid,uname,flowseq,signdate,status,comment,agnname) " &
        "values(" & CLng(TxtAttaDoc.Text) & ",'" & signid & "','" & signname & "'," & flowseq &
        ",'" & signdate & "','" & reason & "','" & comment & "','" & agnname & "')"
        CommUtil.SqlSapExecute("ins", SqlCmd, connsap)
        connsap.Close()
    End Sub


    Sub Email_SignOffFlow(action As String, mtype As Integer) 'mtype :0(normal)  1 :skip function
        Dim body, tdate As String
        Dim now_person, formname, subject, subject1, start_person, last_person, next_person, emailadd, startemail, nowid, startid As String
        Dim sendseq, maxseq As Integer
        Dim infostr As String
        Dim urlpara As String
        Dim statusindex As Integer
        Dim final_singoff As Boolean
        Dim agnid As String
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim signid As String
        Dim signmode As String
        Dim justone As Boolean
        justone = False
        signmode = "signoff"
        emailadd = ""
        If (mtype = 1) Then
            signid = Request.QueryString("skipid")
        Else
            signid = Session("s_id")
        End If

        final_singoff = False
        now_person = "NA"
        next_person = "NA"
        nowid = ""
        SqlCmd = "select max(seq) from  [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        dr.Read()
        maxseq = dr(0)
        dr.Close()
        connsap.Close()
        infostr = "簽核通知"
        tdate = Format(Now(), "yyyy/MM/dd HH:mm:ss")
        SqlCmd = "select T0.subject,T1.sfname from  [dbo].[@XASCH] T0 INNER JOIN [dbo].[@XSFTT] T1 ON T0.sfid=T1.sfid where T0.docnum=" & docnum
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        dr.Read()
        subject = dr(0)
        formname = dr(1)
        dr.Close()
        connsap.Close()
        SqlCmd = "select uname,seq from [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum & " and uid='" & signid & "'"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        dr.Read()
        last_person = dr(0)
        sendseq = dr(1)
        dr.Close()
        connsap.Close()
        SqlCmd = "select uname,emailadd,uid from  [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum & " and seq=1"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        dr.Read()
        start_person = dr(0)
        startemail = dr(1)
        startid = dr(2)
        dr.Close()
        connsap.Close()
        statusindex = 0
        If (action = "send") Then
            last_person = start_person
            SqlCmd = "select uname,emailadd,uid from  [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum & " and seq=2" '@XSPWT (簽核人員Table)
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                dr.Read()
                now_person = dr(0)
                emailadd = dr(1)
                nowid = dr(2)
                dr.Close()
                connsap.Close()
                If (maxseq >= 3) Then
                    SqlCmd = "select uname from  [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum & " and seq=3"
                    dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                    dr.Read()
                    next_person = dr(0)
                    dr.Close()
                    connsap.Close()
                Else
                    next_person = "無"
                End If
            Else
                dr.Close()
                connsap.Close()
                final_singoff = True
                now_person = "結束"
                next_person = "結束"
                justone = True
            End If
        ElseIf (action = "archive") Then

        ElseIf (action = "reject") Then
            '寄email 給送審者
            'signmode = "single" ==> 如不列入連續簽核 , 應設為 single_signoff
            infostr = "退回通知"
            now_person = start_person
            emailadd = startemail
            nowid = startid
            SqlCmd = "select uname from  [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum & " and seq=2"
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            next_person = dr(0)
            dr.Close()
            connsap.Close()
        ElseIf (action = "against") Then
            '寄email 給下一關
            SqlCmd = "select uname,emailadd,uid from  [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum & " and seq=" & sendseq + 1
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            now_person = dr(0)
            emailadd = dr(1)
            nowid = dr(2)
            dr.Close()
            connsap.Close()

            '下一關人員
            If (maxseq > (sendseq + 1)) Then
                SqlCmd = "select uname from  [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum & " and seq=" & sendseq + 2
                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                dr.Read()
                next_person = dr(0)
                dr.Close()
                connsap.Close()
            Else
                next_person = "結束"
            End If
        ElseIf (action = "approval" Or action = "skip") Then
            '寄email 給下一關
            If (maxseq > sendseq) Then
                SqlCmd = "select uname,emailadd,uid from  [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum & " and seq=" & sendseq + 1
                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                dr.Read()
                now_person = dr(0)
                emailadd = dr(1)
                nowid = dr(2)
                dr.Close()
                connsap.Close()
            Else
                final_singoff = True
                now_person = "結束"
                'statusindex = 6
            End If
            '下一關人員
            If (maxseq > (sendseq + 1)) Then
                SqlCmd = "select uname from  [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum & " and seq=" & sendseq + 2
                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                dr.Read()
                next_person = dr(0)
                dr.Close()
                connsap.Close()
            Else
                next_person = "結束"
            End If
        End If
        Dim href As String
        href = url & "usermgm/login.aspx"
        If (final_singoff) Then
            SqlCmd = "select uname,emailadd,uid from  [dbo].[@XSPWT] where signprop=1 and docentry=" & docnum
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            now_person = dr(0)
            emailadd = dr(1)
            nowid = dr(2)
            dr.Close()
            connsap.Close()
            infostr = "歸檔通知"
            'Dim di As DirectoryInfo
            'di = New DirectoryInfo(HttpContext.Current.Server.MapPath("~/") & "AttachFile\SignOffsFormFiles\" & sid_create & "\" & sfid & "\")
            'Dim fi As FileInfo()
            'Dim fname As String
            'fi = di.GetFiles(docnum & "*主檔*")
            '以下為產生簽核條pdf 與原內容pdf 檔案合併
            Dim fileNameSign As String
            Dim pdfFiles(1) As String
            Dim fileNameApproved As String
            fileNameSign = docnum & "_sign.pdf" '簽核條pdf
            'mergePDF(ByVal pdfFiles() As String, ByVal savefileName As String, ByVal fpath As String) As Boolean
            If (sfid <> 100 And sfid <> 101) Then
                fileNameApproved = docnum & "_簽核主檔" & CommSignOff.GetCurrentAttachedFileName(docnum) & "_Approved.pdf"
            Else
                fileNameApproved = docnum & "_加簽核主檔" & CommSignOff.GetCurrentAttachedFileName(docnum) & "_Approved.pdf"
            End If

            If (sfid = 51 Or sfid = 50 Or sfid = 49) Then 'sfid Process '入出庫單 '
                If (Not System.IO.Directory.Exists(targetPath)) Then '因一開始不是用附檔 , 故目錄可能不存在
                    Directory.CreateDirectory(targetPath)
                End If
                If (Not System.IO.Directory.Exists(localsignoffformdir)) Then '因一開始不是用附檔 , 故目錄可能不存在
                    Directory.CreateDirectory(localsignoffformdir)
                End If
                CommSignOff.createMaterialInOutPDF(docnum, formname, fileNameApproved, targetPath, TxtPrice.Text, TxtSubject.Text, "NTD", sfid)
            ElseIf (sfid = 16) Then
                If (Not System.IO.Directory.Exists(targetPath)) Then '因一開始不是用附檔 , 故目錄可能不存在
                    Directory.CreateDirectory(targetPath)
                End If
                If (Not System.IO.Directory.Exists(localsignoffformdir)) Then '因一開始不是用附檔 , 故目錄可能不存在
                    Directory.CreateDirectory(localsignoffformdir)
                End If
                CommSignOff.createRSCPDF(docnum, "門禁磁卡補刷卡單", fileNameApproved, targetPath, sfid)
            ElseIf (sfid = 12 Or sfid = 1 Or sfid = 22 Or sfid = 3 Or sfid = 100 Or sfid = 23 Or sfid = 24 Or sfid = 101) Then 'sfid process 用HTML 轉 PDF 方式
                If (Not System.IO.Directory.Exists(targetPath)) Then '因一開始不是用附檔 , 故目錄可能不存在
                    Directory.CreateDirectory(targetPath)
                End If
                If (Not System.IO.Directory.Exists(localsignoffformdir)) Then '因一開始不是用附檔 , 故目錄可能不存在
                    Directory.CreateDirectory(localsignoffformdir)
                End If
                'CommSignOff.HtmlToPdfGen(url, docnum, sfid, localsapuploaddir) '改用下面 , 因可能會有檔名一樣 , 複製時會因電腦cache問題,用到cache之檔案
                'File.Copy(localsapuploaddir & "gencltemp.pdf", targetPath & fileNameApproved) '將由HtmlT0Pdf產生之pdf copy至共同目錄 , 之後在交由下述Stamper1 處理
                CommSignOff.HtmlToPdfGen(url, docnum, sfid, targetPath & fileNameApproved)

                If (sfid = 100) Then
                    RecordSignFlowHistotySfid100_101("補充加簽單號:" & docnum, "補充加簽完成")
                ElseIf (sfid = 101) Then
                    RecordSignFlowHistotySfid100_101("返還單號:" & docnum, "料件返還完成")
                    '將設定之返還數量加到rtnqty
                    If (TxtAttaDoc.Text <> "NA" And TxtAttaDoc.Text <> "") Then
                        SqlCmd = "update [dbo].[@XSMLS] set rtnqty=rtnqty+nowrtnqty " &
                            " where docentry=" & CLng(TxtAttaDoc.Text) & " and head=0"
                        CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                        connsap.Close()
                    End If
                    '將設定之返還數量歸0
                    If (TxtAttaDoc.Text <> "NA" And TxtAttaDoc.Text <> "") Then
                        SqlCmd = "update [dbo].[@XSMLS] set nowrtnqty=0 " &
                            " where docentry=" & CLng(TxtAttaDoc.Text) & " and head=0"
                        CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                        connsap.Close()
                    End If
                    '更新todo 事項之完成%
                    SqlCmd = "Select sum(quantity),sum(rtnqty) FROM [dbo].[@XSMLS] T0 WHERE T0.[docentry] =" & CLng(TxtAttaDoc.Text) & " and head=0"
                    dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, conn)
                    dr.Read()
                    Dim a As Double
                    Dim b As Integer
                    a = (dr(1) / dr(0)) * 10
                    b = Math.Round(a) * 10
                    dr.Close()
                    conn.Close()
                    SqlCmd = "update dbo.[@XSTDT] set status= " & b & ",upddate='" & Format(Now(), "yyyy/MM/dd") & "' where docentry=" & CLng(TxtAttaDoc.Text)
                    CommUtil.SqlSapExecute("upd", SqlCmd, conn)
                    conn.Close()
                End If
                'MsgBox(targetPath & fileNameApprovedTemp)
            Else '這理是沒內建表單,故簽核條獨立於原附檔內容外(之後會與原主簽核檔合併,其他附檔當附件),而上述之ElseIf 直接產生內容加簽核條pdf(若有附檔,不做合併,當附件)
                If (sfid > 50 And sfid < 80) Then
                    CommSignOff.createProcessDtlWithPricePDF(docnum, formname, fileNameSign, targetPath, TxtPrice.Text, TxtSubject.Text, DDLDollorUnit.SelectedValue)
                Else
                    CommSignOff.createProcessDtlPDF(docnum, formname, fileNameSign, targetPath, TxtPrice.Text, TxtSubject.Text, DDLDollorUnit.SelectedValue)
                End If
            End If
            Dim specialsfid As Integer
            Dim di As DirectoryInfo
            Dim fi As FileInfo()
            Dim fileNameStamped As String
            Dim siddir As String
            Dim sfidnum As Integer
            Dim k As Integer
            di = New DirectoryInfo(HttpContext.Current.Server.MapPath("~/") & "AttachFile\SignOffsFormFiles\" & sid_create & "\" & sfid & "\")
            If (CommSignOff.IsSelfForm(sfid) = 1) Then 'sfid process
                specialsfid = 1
                If (sfid = 100) Then
                    fileNameStamped = docnum & "_加簽核主檔" & CommSignOff.GetCurrentAttachedFileName(docnum) & "_Stamped.pdf"
                ElseIf (sfid = 101) Then
                    fileNameStamped = docnum & "_返還簽核主檔" & CommSignOff.GetCurrentAttachedFileName(docnum) & "_Stamped.pdf"
                Else
                    fileNameStamped = docnum & "_簽核主檔" & CommSignOff.GetCurrentAttachedFileName(docnum) & "_Stamped.pdf"
                End If
            Else
                fi = di.GetFiles(docnum & "*主檔*")
                pdfFiles(0) = fi(0).Name
                pdfFiles(1) = fileNameSign '簽核條pdf
                CommSignOff.mergePDF(pdfFiles, fileNameApproved, targetPath)
                specialsfid = 0
                fileNameStamped = docnum & "_簽核主檔" & CommSignOff.GetCurrentAttachedFileName(docnum) & "_" & Split(fi(0).Name, ".")(0) & "_Stamped.pdf"
            End If
            Threading.Thread.Sleep(1000) '等fileNameApproved穩定
            CommSignOff.GenPdfStamper1(fileNameApproved, fileNameStamped, targetPath, localsignoffformdir, specialsfid, True) '產生浮水印pdf 並刪除approved pdf
            '以下找尋未打上浮水印之pdf 附檔,並打上Jet Logo 浮水印
            fi = di.GetFiles("*附檔*")
            For k = 0 To fi.Length - 1
                If (InStr(fi(k).Name, "Stamped") = 0 And InStr(LCase(fi(k).Name), "pdf") <> 0) Then
                    'MsgBox(fi(k).Name)
                    CommSignOff.GenPdfStamper1(fi(k).Name, Split(fi(k).Name, ".")(0) & "_Stamped.pdf", targetPath, localsignoffformdir, specialsfid, False) '產生浮水印pdf 並刪除approved pdf
                End If
            Next
            If (sfid = 100 Or sfid = 101) Then '將加簽單merge到主單 '
                Dim subformdir, mainformdir As String
                Dim NewStampFileName, mainlocalsignoffformdir As String
                NewStampFileName = ""
                SqlCmd = "Select sid,sfid from [dbo].[@XASCH] where docnum=" & CLng(TxtAttaDoc.Text)
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                drL.Read()
                siddir = drL(0)
                sfidnum = drL(1)
                drL.Close()
                connL.Close()
                subformdir = targetPath
                mainformdir = HttpContext.Current.Server.MapPath("~/") & "AttachFile\SignOffsFormFiles\" & siddir & "\" & sfidnum & "\"
                di = New DirectoryInfo(mainformdir)
                fi = di.GetFiles(CLng(TxtAttaDoc.Text) & "*主檔*")
                pdfFiles(0) = fileNameStamped
                pdfFiles(1) = fi(0).Name
                SqlCmd = "update [dbo].[@XASCH] set attachfileno=attachfileno+1 " &
                        " where docnum=" & CLng(TxtAttaDoc.Text)
                CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                connsap.Close()
                If (sfid = 100) Then
                    NewStampFileName = TxtAttaDoc.Text & "_簽核(加)主檔" & CommSignOff.GetCurrentAttachedFileName(CLng(TxtAttaDoc.Text)) & "_Stamped.pdf"
                ElseIf (sfid = 101) Then
                    NewStampFileName = TxtAttaDoc.Text & "_簽核(返還)主檔" & CommSignOff.GetCurrentAttachedFileName(CLng(TxtAttaDoc.Text)) & "_Stamped.pdf"
                End If
                '將合併檔(NewStampFileName)存到主單目錄
                CommSignOff.mergePDFOfDiffDir(pdfFiles, NewStampFileName, subformdir, mainformdir)
                mainlocalsignoffformdir = Application("localdir") & "SignOffsFormFiles\" & siddir & "\" & sfidnum & "\"
                Threading.Thread.Sleep(1000) '等NewStampFileName穩定
                File.Copy(mainformdir & NewStampFileName, mainlocalsignoffformdir & NewStampFileName)
                '在http程式目錄刪除主單即加簽目錄之主檔
                File.Delete(mainformdir & pdfFiles(1))
                File.Delete(subformdir & pdfFiles(0))
                '在備份目錄刪除主單即加簽目錄之主檔
                File.Delete(mainlocalsignoffformdir & pdfFiles(1))
                File.Delete(localsignoffformdir & pdfFiles(0))
            End If

            '如設定是追蹤單據 , 則需寫入@XSTDT
            Dim todoflag As Integer
            SqlCmd = "Select T0.todoflag from [dbo].[@XSFTT] T0 " & 'XSFTT 簽核表單種類
                "where T0.sfid=" & sfid
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                dr.Read()
                todoflag = dr(0)
            End If
            dr.Close()
            connsap.Close()
            If (todoflag = 1) Then
                SqlCmd = "insert into [dbo].[@XSTDT] (docentry,sfid,CDate,subject,traceperson) " &
                        "values(" & docnum & "," & sfid & ",'" & tdate & "','" & subject & "','" & nowid & "')"
                CommUtil.SqlSapExecute("ins", SqlCmd, connsap)
                connsap.Close()
            End If
        End If
        urlpara = "?actmode=" & signmode & "&uid=" & nowid & "&status=" & docstatus & "&formtypeindex=" & formtypeindex &
                "&formstatusindex=" & statusindex & "&docnum=" & docnum & "&sfid=" & sfid
        body = "<span><h5>此信件為系統發出信件，請勿直接回覆，感謝您的配合!</h5></span>" &
                "<table border=1 width=360 border-collapse:collapse>" &
                "<tr bgcolor=#add8e6><td><font size=6><b>" & infostr & "</b></font></h1>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;待核決人&nbsp;:&nbsp;" & now_person & "</td></tr>" &
                "<tr><td align=center><a href=" & href & urlpara & ">前往簽核系統</a></td></tr>" &
                "<tr><td>單據編號&nbsp;:&nbsp;" & docnum & "</td></tr>" &
                "<tr><td>單據名稱&nbsp;:&nbsp;" & formname & "</td></tr>" &
                "<tr><td>主旨&nbsp;:&nbsp;" & subject & "</td></tr>" &
                "<tr><td>送審人&nbsp;:&nbsp;" & start_person & "</td></tr>" &
                "<tr><td>上一關&nbsp;:&nbsp;" & last_person & "</td></tr>" &
                "<tr><td>目前關&nbsp;:&nbsp;" & now_person & "</td></tr>" &
                "<tr><td>下一關&nbsp;:&nbsp;" & next_person & "</td></tr>" &
                "<tr><td>通知日期&nbsp;:&nbsp;" & tdate & "</td></tr>" &
                "</table>"
        subject1 = "(捷智簽核) " & infostr & ":" & formname & " - " & subject
        'emailadd = "ron@jettech.com.tw" 'temp test
        CommUtil.SendMail(emailadd, subject1, body)
        agnid = CommSignOff.AgencySet(nowid)
        If (agnid <> "") Then '啟動代理人郵件通知
            SqlCmd = "select name,email from dbo.[user] where id='" & agnid & "'"
            drL = CommUtil.SelectLocalSqlUsingDr(SqlCmd, connL)
            If (drL.HasRows) Then
                drL.Read()
                infostr = drL(0) & " 代理簽核通知"
                urlpara = "?actmode=signoff&uid=" & nowid & "&status=" & docstatus & "&formtypeindex=" & formtypeindex &
                "&formstatusindex=" & statusindex & "&docnum=" & docnum & "&sfid=" & sfid & "&agnid=" & agnid
                body = "<span><h5>此信件為系統發出信件，請勿直接回覆，感謝您的配合!</h5></span>" &
                "<table border=1 width=360 border-collapse:collapse>" &
                "<tr bgcolor=#add8e6><td><font size=6><b>" & infostr & "</b></font></h1>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;待核決人&nbsp;:&nbsp;" & now_person & "</td></tr>" &
                "<tr><td align=center><a href=" & href & urlpara & ">前往簽核系統</a></td></tr>" &
                "<tr><td>單據編號&nbsp;:&nbsp;" & docnum & "</td></tr>" &
                "<tr><td>單據名稱&nbsp;:&nbsp;" & formname & "</td></tr>" &
                "<tr><td>主旨&nbsp;:&nbsp;" & subject & "</td></tr>" &
                "<tr><td>送審人&nbsp;:&nbsp;" & start_person & "</td></tr>" &
                "<tr><td>上一關&nbsp;:&nbsp;" & last_person & "</td></tr>" &
                "<tr><td>目前關(代理人)&nbsp;:&nbsp;" & now_person & "(" & drL(0) & ")</td></tr>" &
                "<tr><td>下一關&nbsp;:&nbsp;" & next_person & "</td></tr>" &
                "<tr><td>通知日期&nbsp;:&nbsp;" & tdate & "</td></tr>" &
                "</table>"
                subject1 = "(捷智簽核) " & infostr & ":" & formname & " - " & subject
                'emailadd = "ron@jettech.com.tw" 'temp test
                CommUtil.SendMail(emailadd, subject1, body)
            Else
                CommUtil.ShowMsg(Me, "無法在User資料表中找到" & agnid & "代理人資料")
            End If
            drL.Close()
            connL.Close()
        End If
        If (final_singoff And nowid <> startid And justone = False) Then
            SqlCmd = "select uname,emailadd,uid from  [dbo].[@XSPWT] where signprop=1 and docentry=" & docnum
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            now_person = start_person
            emailadd = startemail
            nowid = startid
            dr.Close()
            connsap.Close()
            infostr = "簽核結束通知"
            urlpara = "?actmode=signoff&uid=" & nowid & "&status=" & docstatus & "&formtypeindex=" & formtypeindex &
                "&formstatusindex=" & statusindex & "&docnum=" & docnum & "&sfid=" & sfid
            body = "<span><h5>此信件為系統發出信件，請勿直接回覆，感謝您的配合!</h5></span>" &
                "<table border=1 width=360 border-collapse:collapse>" &
                "<tr bgcolor=#add8e6><td><font size=6><b>" & infostr & "</b></font></h1>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;被通知人&nbsp;:&nbsp;" & now_person & "</td></tr>" &
                "<tr><td align=center><a href=" & href & urlpara & ">前往簽核系統</a></td></tr>" &
                "<tr><td>單據編號&nbsp;:&nbsp;" & docnum & "</td></tr>" &
                "<tr><td>單據名稱&nbsp;:&nbsp;" & formname & "</td></tr>" &
                "<tr><td>主旨&nbsp;:&nbsp;" & subject & "</td></tr>" &
                "<tr><td>送審人&nbsp;:&nbsp;" & start_person & "</td></tr>" &
                "<tr><td>上一關&nbsp;:&nbsp;" & last_person & "</td></tr>" &
                "<tr><td>目前關&nbsp;:&nbsp;" & now_person & "</td></tr>" &
                "<tr><td>下一關&nbsp;:&nbsp;" & next_person & "</td></tr>" &
                "<tr><td>通知時間&nbsp;:&nbsp;" & tdate & "</td></tr>" &
                "</table>"
            subject1 = "(捷智簽核) " & infostr & ":" & formname & " - " & subject
            'emailadd = "ron@jettech.com.tw" 'temp test
            CommUtil.SendMail(emailadd, subject1, body)
        End If

        If (final_singoff) Then
            SqlCmd = "select uname,emailadd,uid from  [dbo].[@XSPWT] where signprop=2 and docentry=" & docnum
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                Do While (dr.Read())
                    now_person = dr(0)
                    emailadd = dr(1)
                    nowid = dr(2)
                    infostr = "事件知悉通知"
                    urlpara = "?actmode=signoff&uid=" & nowid & "&status=" & docstatus & "&formtypeindex=" & formtypeindex &
                    "&formstatusindex=" & statusindex & "&docnum=" & docnum & "&sfid=" & sfid
                    body = "<span><h5>此信件為系統發出信件，請勿直接回覆，感謝您的配合!</h5></span>" &
                    "<table border=1 width=360 border-collapse:collapse>" &
                    "<tr bgcolor=#add8e6><td><font size=6><b>" & infostr & "</b></font></h1>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;待知悉人&nbsp;:&nbsp;" & now_person & "</td></tr>" &
                    "<tr><td align=center><a href=" & href & urlpara & ">前往簽核系統</a></td></tr>" &
                    "<tr><td>單據編號&nbsp;:&nbsp;" & docnum & "</td></tr>" &
                    "<tr><td>單據名稱&nbsp;:&nbsp;" & formname & "</td></tr>" &
                    "<tr><td>主旨&nbsp;:&nbsp;" & subject & "</td></tr>" &
                    "<tr><td>送審人&nbsp;:&nbsp;" & start_person & "</td></tr>" &
                    "<tr><td>上一關&nbsp;:&nbsp;" & last_person & "</td></tr>" &
                    "<tr><td>目前關&nbsp;:&nbsp;" & now_person & "</td></tr>" &
                    "<tr><td>下一關&nbsp;:&nbsp;" & next_person & "</td></tr>" &
                    "<tr><td>通知日期&nbsp;:&nbsp;" & tdate & "</td></tr>" &
                    "</table>"
                    subject1 = "(捷智簽核) " & infostr & ":" & formname & " - " & subject
                    'emailadd = "ron@jettech.com.tw" 'temp test
                    CommUtil.SendMail(emailadd, subject1, body)
                Loop
            End If
            dr.Close()
            connsap.Close()
        End If
    End Sub
    Function GetSinoffList()
        Dim formstatus As Integer
        formstatus = 1
        'SqlCmd = "SELECT tprice=0,unit='NA',T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate,signoffflag=0 " &
        '             "FROM dbo.[@XASCH] T1 " &
        '             "where T1.sid='" & sid & "' and (T1.status='E' or T1.status='D' or T1.status='B' or T1.status='R') " &
        '             "order by T1.sfid,T1.docnum desc"
        '    ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1) '未送審
        '    connsap1.Close()
        'SqlCmd = "Select tprice=0,unit='NA',T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate,T0.seq,signoffflag=0 " &
        '     "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum " &
        '     "where T0.signprop=0 and T0.status=1 and T1.status<>'B' and T1.status<>'R' and T0.uid='" & sid & "' " &
        '     " order by T0.signprop,T1.sfid,T1.docnum desc"
        '    ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1) '關卡簽核 
        'connsap1.Close()
        SqlCmd = "Select tprice=0,unit='NA',T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate,T0.seq,signoffflag=0 " &
             "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum " &
             "where T0.signprop=0 and T0.status=1 and T0.uid='" & sid & "' " &
             " order by T0.signprop,T1.sfid,T1.docnum desc"
        ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1) '關卡簽核 (把退回及抽回之起始人亦納入)
        connsap1.Close()
        SqlCmd = "Select tprice=0,unit='NA',T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate,T0.seq,signoffflag=0 " &
            "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum " &
            "where T0.signprop=1 And T0.status=1 and T0.uid='" & sid & "' " &
            " order by T0.signprop,T1.sfid,T1.docnum desc"
        ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1) '待歸檔
        connsap1.Close()
        If (agnidG = "") Then '如果是代理人來簽核 , 待知悉不需代理(agnidG 不是空白 , 表示是代理人簽核)
            SqlCmd = "Select T1.price,T1.priceunit,T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate,T0.seq,signoffflag=0 " &
                 "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum " &
                 "where T0.signprop=2 And T0.status=1 and T0.uid ='" & sid & "' " &
                 " order by T1.sfid,T1.docnum desc"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1) '待知悉
            connsap1.Close()
            'MsgBox(ds.Tables(0).Rows(Session("startindex"))("docnum"))
        End If
        Dim row As Integer
        Dim matchdoc As Boolean
        Dim actionstatus, nowformstatus As String
        Dim mes(2) As String
        matchdoc = False
        actionstatus = ""
        nowformstatus = ""
        Session("sgcount") = ds.Tables(0).Rows.Count
        If (docstatus <> "E" And docstatus <> "D") Then
            For row = 0 To ds.Tables(0).Rows.Count - 1
                If (CLng(ds.Tables(0).Rows(row)("docnum")) = docnum) Then
                    matchdoc = True
                    Exit For
                End If
            Next
        Else
            matchdoc = True
        End If
        signoffalready = False
        If (matchdoc = False) Then
            mes = CommSignOff.FormStatusMes(docnum, docstatus)
            info = mes(1)
            docstatus = mes(0)
            '底下敘述以改由上述替代
            'SqlCmd = "SELECT status " &
            '                "FROM dbo.[@XSPHT] where uid='" & Session("s_id") & "' and docentry=" & docnum & " order by flowseq desc"
            'dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap1)
            'If (dr.HasRows) Then
            '    dr.Read()
            '    actionstatus = dr(0)
            'Else
            '    If (docstatus <> "E" And docstatus <> "D") Then
            '        actionstatus = "刪除"
            '    End If
            'End If
            'dr.Close()
            'connsap1.Close()
            'SqlCmd = "SELECT status " &
            '                "FROM dbo.[@XASCH] where docnum=" & docnum
            'dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap1)
            'If (dr.HasRows) Then
            '    dr.Read()
            '    If (dr(0) = "O") Then
            '        nowformstatus = "且此單目前狀態還在簽核中"
            '    ElseIf (dr(0) = "F") Then
            '        nowformstatus = "且此單目前狀態已簽核結案"
            '    ElseIf (dr(0) = "C") Then
            '        nowformstatus = "且此單目前狀態已被作廢"
            '    ElseIf (dr(0) = "R") Then
            '        nowformstatus = "且此單目前狀態被抽回"
            '    ElseIf (dr(0) = "B") Then
            '        nowformstatus = "且此單目前狀態已被退回"
            '    ElseIf (dr(0) = "T") Then
            '        nowformstatus = "且此單目前狀態已簽核結案並歸檔"
            '        docstatus = "T" '當其用email 之link 點選時 , 若之前已被歸檔過 , 因之前docstatus=F , 此處要改為T , 才能使歸檔button隱藏
            '    End If
            'Else
            '    nowformstatus = "且此單目前狀態被刪除"
            'End If
            'dr.Close()
            'connsap1.Close()
            ''row = 1000 '
            'If (actionstatus <> "") Then
            '    If (actionstatus <> "跳過簽核") Then
            '        info = "此筆簽核資料曾被您(或代理人) '" & actionstatus & "' 覆核過," & nowformstatus
            '    Else
            '        info = "此筆簽核資料已被管理者 '" & actionstatus & "'," & nowformstatus
            '    End If
            'Else
            '    info = "此筆簽核資料已被前一人抽單," & nowformstatus
            'End If
            signoffalready = True
            'Response.Redirect("~/invalid.aspx?info=" & info)
        End If
        Return row
    End Function
    Protected Sub BtnPdf_Click(sender As Object, e As EventArgs)
        Dim p, p1 As New Process()

        p.StartInfo.FileName = "C:\program files\wkhtmltopdf\bin\wkhtmltopdf.exe"
        'p.StartInfo.Arguments = url & "signoff/printform.aspx?docnum=" & docnum & " " & Application("localdir") & "gencltemp.pdf"
        p.StartInfo.Arguments = url & "signoff/cLsignoff.aspx?sfid=12&docnum=" & docnum & " " & Application("localdir") & "gencltemp.pdf"
        p.StartInfo.WindowStyle = ProcessWindowStyle.Maximized 'WindowStyle可以設定開啟視窗的大小
        p.StartInfo.Verb = "open"
        p.StartInfo.CreateNoWindow = False
        p.Start()
        p.WaitForExit(5000)
        p.Close()
        p.Dispose()
        'p1.StartInfo.FileName = Application("localdir") & "copyfile.bat"
        'p1.StartInfo.Arguments = " " & Application("localdir") & "gencltemp.pdf " & HttpContext.Current.Server.MapPath("~/") & "TempFile\gencltemp.pdf"
        'p1.StartInfo.WindowStyle = ProcessWindowStyle.Maximized 'WindowStyle可以設定開啟視窗的大小
        'p1.Start()
        'p1.WaitForExit(3000)
        'p1.Close()
        'p1.Dispose()
        'Dim tpath As String = url & "TempFile/gencltemp.pdf"
        'ScriptManager.RegisterStartupScript(Me, Page.GetType, "Script", "showDisplay('" & tpath & "');", True)

        'Dim screenWidth As Integer = Windows.Forms.Screen.PrimaryScreen.Bounds.Width
        'Dim screenHeight As Integer = Windows.Forms.Screen.PrimaryScreen.Bounds.Height
        'MsgBox("螢幕解析度為 " + screenWidth.ToString() + "*" + screenHeight.ToString())
        'Dim workarea_Hight As Integer
        'Dim workerarea_width As Integer
        'workarea_Hight = Windows.Forms.Screen.PrimaryScreen.WorkingArea.Width
        'workerarea_width = Windows.Forms.Screen.PrimaryScreen.WorkingArea.Height
        'MsgBox("工作區域大小" & workerarea_width & "X" & workarea_Hight)
        'Dim height As Double = SystemParameters.VirtualScreenHeight
        'Dim width As Double = SystemParameters.VirtualScreenWidth
        'Dim resolution As Double = height * width
        'MsgBox(width & "*" & height)
        'Dim mDSize As Double = Math.Sqrt(Math.Pow(moniPhySize.Width, 2) + Math.Pow(moniPhySize.Height, 2)) / 2.54D
        'MsgBox(moniPhySize.Width & "*" & moniPhySize.Height)
        'Dim oForm As System.Windows.Forms.Form
        'Dim oGraph As System.Drawing.Graphics

        'oGraph = oForm.CreateGraphics()
        'MsgBox(oGraph.DpiX)
        'Dim dpiX, dpiY As Double

        'Dim graphics As System.Drawing.Graphics = this.CreateGraphics

        'dpiX = graphics.DpiX
        'dpiY = graphics.DpiY

        'If (Environment.OSVersion.Version.Major >= 6) Then 'Vista on up
        '    SetProcessDPIAware()
        'End If
        'ScriptManager.RegisterStartupScript(Me, Page.GetType, "Script", "showDisplay();", True)
        'Dim strFilePath As String = "c:\data\kk.pdf"
        'Dim dwl As New System.Net.WebClient()

        'dwl.DownloadFile(tpath, strFilePath)
        'dwl.Dispose()


        'System.Diagnostics.Process.Start("C:\program files\wkhtmltopdf\bin\wkhtmltopdf.exe", "http://localhost:50601/signoff/printform.aspx?docnum=" & docnum & " C:\data\myFileName.pdf")
        'System.Diagnostics.Process.Start("C:\program files\wkhtmltopdf\bin\wkhtmltopdf.exe", "http://192.168.100.109:8080/signoff/printform.aspx?docnum=" & docnum & " C:\data\myFileName.pdf")
        'Thread.Sleep(1000)
        'Process.Start("path=C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader")
        'System.Diagnostics.Process.Start("C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32", "C:\data\myFileName.pdf")
        'Dim p As New Process()
        'p.StartInfo.FileName = "AcroRd32.exe"
        'p.StartInfo.Arguments = "C:\data\myFileName.pdf"
        'p.StartInfo.WindowStyle = ProcessWindowStyle.Maximized 'WindowStyle可以設定開啟視窗的大小
        'p.StartInfo.Verb = "open"
        'p.StartInfo.CreateNoWindow = False
        'p.Start()
        'p.WaitForExit(5000)
        'If (Not p.HasExited) Then
        '    '測試處理序是否還有回應
        '    If (p.Responding) Then
        '        '關閉使用者介面的處理序
        '        p.CloseMainWindow()
        '    Else
        '        '立即停止相關處理序。意即，處理序沒回應，強制關閉
        '        p.Kill()
        '    End If
        'End If
        'p.Close()
        'p.Dispose()
        'Response.Redirect("~/signoff/printform.aspx?docnum=" & docnum)
    End Sub
    Sub ConvertASPXtoHTML()
        Dim targetASPX As String = "http://localhost:50601/signoff/printform.aspx?docnum=" & docnum 'ConfigurationManager.AppSettings("targetASPX")
        Dim savePath As String = "C:\data\test.html" 'ConfigurationManager.AppSettings("savePath")
        Dim wc As WebClient = New WebClient()
        wc.Encoding = System.Text.Encoding.UTF8
        Dim html As String = wc.DownloadString(targetASPX)
        'If (FileStream fs = New FileStream(savePath, System.IO.FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))Then
        Dim fs As FileStream = New FileStream(savePath, System.IO.FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite)
        Dim sw As StreamWriter = New StreamWriter(fs)
        '輸出html字串
        sw.Write(html)
        sw.Close()
        fs.Close()

        'End If
    End Sub
    Protected Sub ChkDel_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim ftnum As String
        Dim str() As String
        Dim ChkDel As New CheckBox
        Dim BtnFileAct As New Button
        Dim delflag As Boolean
        delflag = False
        'If (DDLAttaFile.SelectedIndex <> 0) Then
        str = Split(sender.ID, "_")
        ftnum = str(2)
        ChkDel = CType(FT_m.FindControl("chk_del_" & ftnum), CheckBox)
        BtnFileAct = CType(FT_m.FindControl("btn_fileact_" & ftnum), Button)
        If (ChkDel.Checked) Then
            If (DDLAttaFile.SelectedIndex <> 0) Then
                str = Split(DDLAttaFile.SelectedValue, "_")
                If (CLng(str(0)) = docnum) Then
                    BtnFileAct.Text = "刪除"
                Else
                    ChkDel.Checked = False
                    CommUtil.ShowMsg(Me, "此選擇之檔案不屬於此單號,不能刪除")
                End If
            Else
                CommUtil.ShowMsg(Me, "尚未選擇刪除檔案")
            End If
        Else
            BtnFileAct.Text = "上傳"
        End If
        'Else
        'ChkDel.Checked = False
        'CommUtil.ShowMsg(Me, "需先選擇刪除檔案")
        'End If
    End Sub
    Protected Sub ChkReturn_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim ftnum As String
        Dim str() As String
        Dim ChkR As CheckBox
        str = Split(sender.ID, "_")
        ftnum = str(2)
        If (ftnum = "0") Then
            ChkR = CType(FT_0.FindControl("chk_return_" & ftnum), CheckBox)
            If (ChkR.Checked) Then
                CType(FT_m.FindControl("chk_return_m"), CheckBox).Checked = True
                CType(FT_1.FindControl("chk_return_1"), CheckBox).Checked = True
            Else
                CType(FT_m.FindControl("chk_return_m"), CheckBox).Checked = False
                CType(FT_1.FindControl("chk_return_1"), CheckBox).Checked = False
            End If
        ElseIf (ftnum = "m") Then
            ChkR = CType(FT_m.FindControl("chk_return_" & ftnum), CheckBox)
            If (ChkR.Checked) Then
                CType(FT_0.FindControl("chk_return_0"), CheckBox).Checked = True
                CType(FT_1.FindControl("chk_return_1"), CheckBox).Checked = True
            Else
                CType(FT_0.FindControl("chk_return_0"), CheckBox).Checked = False
                CType(FT_1.FindControl("chk_return_1"), CheckBox).Checked = False
            End If
        ElseIf (ftnum = "1") Then
            ChkR = CType(FT_1.FindControl("chk_return_" & ftnum), CheckBox)
            If (ChkR.Checked) Then
                CType(FT_0.FindControl("chk_return_0"), CheckBox).Checked = True
                CType(FT_m.FindControl("chk_return_m"), CheckBox).Checked = True
            Else
                CType(FT_0.FindControl("chk_return_0"), CheckBox).Checked = False
                CType(FT_m.FindControl("chk_return_m"), CheckBox).Checked = False
            End If
        End If

    End Sub
    Function GetNextAttachedFileName()
        Dim nextfilename As String
        Dim connsap As New SqlConnection
        Dim dr As SqlDataReader
        nextfilename = ""
        SqlCmd = "Select attachfileno " &
        "from [dbo].[@XASCH] " &
        "where docnum=" & docnum
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            dr.Read()
            nextfilename = "(" & CStr(dr(0) + 1) & ")"
        End If
        dr.Close()
        connsap.Close()
        Return nextfilename
    End Function

    Protected Sub BtnFileAct_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim targetDir As String
        'Dim appPath As String
        Dim nameext, attachfile As String
        Dim str() As String
        Dim p As New Process()
        Dim ftnum As String
        Dim FileUL As New FileUpload
        Dim file_exist As Boolean
        Dim act As String
        Dim di As DirectoryInfo
        act = ""
        file_exist = False
        targetDir = HttpContext.Current.Server.MapPath("~/") & "\AttachFile\SignOffsFormFiles\" & Session("s_id") & "\" & sfid & "\"
        'appPath = Request.PhysicalApplicationPath '應用程式目錄
        If (Not System.IO.Directory.Exists(targetDir)) Then
            Directory.CreateDirectory(targetDir)
        End If
        If (Not System.IO.Directory.Exists(localsignoffformdir)) Then
            Directory.CreateDirectory(localsignoffformdir)
        End If
        If File.Exists(targetFile) Then
            file_exist = True
        End If
        str = Split(sender.ID, "_")
        ftnum = str(2)
        If (ftnum = "m") Then
            FileUL = CType(FT_m.FindControl("fileul_m"), FileUpload)
        ElseIf (ftnum = "0") Then
            FileUL = CType(FT_0.FindControl("fileul_0"), FileUpload)
        ElseIf (ftnum = "1") Then
            FileUL = CType(FT_1.FindControl("fileul_1"), FileUpload)
        End If

        If (sender.Text = "上傳") Then
            If (FileUL.HasFile) Then
                str = Split(FileUL.FileName, ".")
                nameext = str(1)
                di = New DirectoryInfo(HttpContext.Current.Server.MapPath("~/") & "AttachFile\SignOffsFormFiles\" & sid_create & "\" & sfid & "\")
                Dim fi As FileInfo()
                Dim fname As String
                If (CType(FT_m.FindControl("chk_subfile"), CheckBox).Checked = False) Then
                    If (nameext <> "pdf") Then
                        CommUtil.ShowMsg(Me, "要上傳之簽核主檔案需為pdf")
                        Exit Sub
                    End If
                    fi = di.GetFiles(docnum & "*主檔*")
                    If (fi.Length <> 0) Then
                        CommUtil.ShowMsg(Me, "已有簽核主檔存在,如要更換,請先刪除原主檔")
                        Exit Sub
                    End If
                    fname = docnum & "_簽核主檔" & GetNextAttachedFileName() & "_" & FileUL.FileName
                Else
                    fi = di.GetFiles(docnum & "*")
                    For i = 0 To fi.Length - 1
                        If (InStr(fi(i).Name, str(0)) <> 0) Then
                            If (InStr(fi(i).Name, "主") <> 0) Then
                                CommUtil.ShowMsg(Me, "欲上傳之附檔,與已存在之簽核主檔同名,請確認是要更新主檔,或選錯檔")
                            Else
                                CommUtil.ShowMsg(Me, "欲上傳之附檔,與已存在之簽核附檔同名,若欲取代,請刪除原附檔,若是新的,請更名")
                            End If
                            Exit Sub
                        End If
                    Next
                    If (sfid = 100) Then
                        fname = docnum & "_加簽核補充附檔" & GetNextAttachedFileName() & "_" & FileUL.FileName
                    ElseIf (sfid = 101) Then
                        fname = docnum & "_返還簽核附檔" & GetNextAttachedFileName() & "_" & FileUL.FileName
                    Else
                        fname = docnum & "_簽核附檔" & GetNextAttachedFileName() & "_" & FileUL.FileName
                    End If
                End If
                FileUL.SaveAs(targetDir & fname)
                FileUL.SaveAs(localsignoffformdir & fname)
                act = "fileup"
                SqlCmd = "update [dbo].[@XASCH] set attachfileno=attachfileno+1 " &
                             " where docnum=" & docnum
                CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                connsap.Close()
            Else
                act = "filenoassign"
            End If
        ElseIf (sender.Text = "刪除") Then
            attachfile = targetDir & DDLAttaFile.SelectedValue
            IO.File.Delete(attachfile)
            IO.File.Delete(localsignoffformdir & DDLAttaFile.SelectedValue)
            sender.Text = "上傳"
            If (ftnum = "m") Then
                CType(FT_m.FindControl("chk_del_" & ftnum), CheckBox).Checked = False
            ElseIf (ftnum = "0") Then
                CType(FT_0.FindControl("chk_del_" & ftnum), CheckBox).Checked = False
            ElseIf (ftnum = "1") Then
                CType(FT_1.FindControl("chk_del_" & ftnum), CheckBox).Checked = False
            End If
            sender.BackColor = Nothing
            'CommUtil.ShowMsg(Me, "刪除成功")
            act = "filedel"
        End If
        'If (sfid > 50 And sfid < 80 And (docstatus = "E" Or docstatus = "A")) Then
        Dim StrPrice As String
        StrPrice = TxtPrice.Text 'Replace(TxtPrice.Text, ",", "")
        If (sfid > 50 And sfid < 80) Then
            Response.Redirect("cLsignoff.aspx?smid=sg&smode=2&act=" & act & "&status=" & docstatus &
                                "&docnum=" & docnum & "&formstatusindex=" & formstatusindex &
                                "&formtypeindex=" & formtypeindex & "&sfid=" & sfid & "&subject=" & TxtSubject.Text &
                                "&sapno=" & TxtSapNO.Text & "&price=" & CDbl(StrPrice) & "&unitindex=" & DDLDollorUnit.SelectedIndex & "&signflowmode=" & signflowmode)
        Else
            Response.Redirect("cLsignoff.aspx?smid=sg&smode=2&act=" & act & "&status=" & docstatus &
                    "&docnum=" & docnum & "&formstatusindex=" & formstatusindex &
                    "&formtypeindex=" & formtypeindex & "&sfid=" & sfid & "&subject=" & TxtSubject.Text & "&signflowmode=" & signflowmode)
        End If
    End Sub

    'Function CellSet(text As String, rowspan As Integer, colspan As Integer, FondBold As Boolean, width As Integer, height As Integer, align As String, BColor As Drawing.Color)
    '    Dim tCell As TableCell
    '    tCell = New TableCell()
    '    tCell.BorderWidth = 1
    '    If (align = "right") Then
    '        tCell.HorizontalAlign = HorizontalAlign.Right
    '    ElseIf (align = "center") Then
    '        tCell.HorizontalAlign = HorizontalAlign.Center
    '    Else
    '        tCell.HorizontalAlign = HorizontalAlign.Left
    '    End If
    '    'If (color) Then
    '    tCell.BackColor = BColor
    '    'End If
    '    tCell.Wrap = True
    '    If (text <> "") Then
    '        tCell.Text = text
    '    End If
    '    tCell.ColumnSpan = colspan
    '    tCell.RowSpan = rowspan
    '    If (width <> 0) Then
    '        tCell.Width = width 'stdwidth * colspan * 0.95
    '    End If
    '    If (height <> 0) Then
    '        tCell.Height = height '20 * rowspan
    '    End If
    '    tCell.Font.Bold = FondBold
    '    Return tCell
    'End Function

    'Function CellSetWithTextBox(rowspan As Integer, colspan As Integer, txtid As String, multiline As Integer, fonesize As Integer, width As Integer, BColor As Drawing.Color)
    '    '如果textbox要以預設大小適合 Cell , 則width設為0 , 如果要調大小 , try 看數值多少
    '    Dim tCell As TableCell
    '    Dim tTxt As New TextBox
    '    tCell = New TableCell()
    '    tCell.Wrap = True
    '    tCell.BorderWidth = 1
    '    tCell.HorizontalAlign = HorizontalAlign.Center
    '    tCell.ColumnSpan = colspan
    '    tCell.RowSpan = rowspan
    '    tTxt.ID = txtid
    '    tTxt.BackColor = BColor
    '    If (fonesize <> 0) Then
    '        tTxt.Font.Size = fonesize
    '    End If
    '    'tTxt.Width = tCell.Width 'stdwidth * colspan * 0.95
    '    'tTxt.Height = tCell.Height '20 * rowspan
    '    If (width <> 0) Then
    '        tTxt.Width = width
    '    End If
    '    If (multiline <> 0) Then
    '        tTxt.TextMode = TextBoxMode.MultiLine
    '        tTxt.Rows = multiline
    '    End If
    '    tCell.Controls.Add(tTxt)
    '    Return tCell
    'End Function
    'Function CellSetWithCalenderExtender(rowspan As Integer, colspan As Integer, txtid As String, ceid As String, BColor As Drawing.Color)
    '    Dim tCell As New TableCell
    '    Dim ce As New CalendarExtender
    '    Dim tTxt As New TextBox
    '    tCell.ColumnSpan = colspan
    '    tCell.RowSpan = rowspan
    '    tCell.BorderWidth = 1
    '    tCell.HorizontalAlign = HorizontalAlign.Center
    '    tTxt.ID = txtid
    '    tTxt.Width = tCell.Width
    '    tTxt.Height = tCell.Height
    '    tTxt.BackColor = BColor
    '    tCell.Controls.Add(tTxt)
    '    ce.TargetControlID = txtid
    '    ce.ID = ceid
    '    ce.Format = "yyyy/MM/dd"
    '    tCell.Controls.Add(ce)
    '    Return tCell
    'End Function
    Function CellSetWithExtender(rowspan As Integer, colspan As Integer, LBxid As String, txtid As String, ddeid As String, BColor As Drawing.Color)
        Dim tCell As New TableCell
        Dim dde As New DropDownExtender
        Dim tTxt As New TextBox
        Dim LBx As New ListBox
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.ColumnSpan = colspan
        tCell.RowSpan = rowspan
        LBx.ID = LBxid
        LBx.AutoPostBack = True
        LBx.Rows = 30
        AddHandler LBx.SelectedIndexChanged, AddressOf LB_SelectedIndexChanged
        tCell.Controls.Add(LBx)
        tTxt.ID = txtid
        tTxt.BackColor = BColor
        'tTxt.Width = 300
        tCell.Controls.Add(tTxt)
        dde.TargetControlID = txtid
        dde.ID = ddeid
        dde.DropDownControlID = LBxid
        tCell.Controls.Add(dde)
        Return tCell
    End Function

    Protected Sub LB_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim tTxt As TextBox
        Dim str() As String
        Dim id As String
        'Dim LBx As ListBox
        'Dim model, mdesc, mtype As String
        str = Split(sender.ID, "_")
        id = str(1)
        tTxt = ContentT.FindControl("txt_" & id)
        str = Split(sender.SelectedValue, "-")
        If (id <> "model") Then
            tTxt.Text = sender.SelectedValue
        Else
            tTxt.Text = str(1)
        End If
    End Sub
    Sub WriteListBoxItemForXCMRT()
        Dim LBx As ListBox
        Dim model, mdesc, mtype As String
        SqlCmd = "SELECT T1.[Name] FROM dbo.[OSCT] T1 ORDER BY T1.Name"
        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (drsap.HasRows) Then
            LBx = ContentT.FindControl("lb_machinetype")
            LBx.Items.Clear()
            LBx.Items.Add("")
            Do While (drsap.Read())
                LBx.Items.Add(drsap(0))
            Loop
        End If
        drsap.Close()
        connsap.Close()

        SqlCmd = "SELECT T0.u_model,T0.u_mdesc,T0.u_mtype " &
                    "FROM dbo.[@UMMD] T0 order by T0.u_model,T0.u_mcode"
        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        LBx = ContentT.FindControl("lb_model")
        LBx.Items.Clear()
        LBx.Items.Add("")
        If (drsap.HasRows) Then
            Do While (drsap.Read())
                model = drsap(0)
                mdesc = drsap(1)
                mtype = drsap(2)
                LBx.Items.Add(mtype & "-" & model & "-" & mdesc)
            Loop
        End If
        LBx.SelectedIndex = 0
        drsap.Close()
        connsap.Close()

        SqlCmd = "SELECT distinct T0.cusname " &
                "FROM dbo.[@XCMRT] T0"
        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (drsap.HasRows) Then
            LBx = ContentT.FindControl("lb_cusname")
            LBx.Items.Clear()
            LBx.Items.Add("")
            Do While (drsap.Read())
                LBx.Items.Add(drsap(0))
            Loop
        End If
        drsap.Close()
        connsap.Close()
        LBx.SelectedIndex = 0

        If (sfid = 22) Then
            SqlCmd = "SELECT distinct T0.cusfactoryOrmo " &
                "FROM dbo.[@XCMRT] T0"
            drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (drsap.HasRows) Then
                LBx = ContentT.FindControl("lb_cusfactoryOrmo")
                LBx.Items.Clear()
                LBx.Items.Add("")
                Do While (drsap.Read())
                    LBx.Items.Add(drsap(0))
                Loop
            End If
            drsap.Close()
            connsap.Close()
            LBx.SelectedIndex = 0

            SqlCmd = "SELECT distinct T0.faeperson " &
                "FROM dbo.[@XCMRT] T0"
            drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (drsap.HasRows) Then
                LBx = ContentT.FindControl("lb_faeperson")
                LBx.Items.Clear()
                LBx.Items.Add("")
                Do While (drsap.Read())
                    LBx.Items.Add(drsap(0))
                Loop
            End If
            drsap.Close()
            connsap.Close()
            LBx.SelectedIndex = 0

            SqlCmd = "SELECT distinct T0.verandspec " &
                "FROM dbo.[@XCMRT] T0"
            drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (drsap.HasRows) Then
                LBx = ContentT.FindControl("lb_verandspec")
                LBx.Items.Clear()
                LBx.Items.Add("")
                Do While (drsap.Read())
                    LBx.Items.Add(drsap(0))
                Loop
            End If
            drsap.Close()
            connsap.Close()
            LBx.SelectedIndex = 0
        ElseIf (sfid = 3) Then
            SqlCmd = "SELECT distinct T0.qcperson " &
                "FROM dbo.[@XCMRT] T0"
            drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (drsap.HasRows) Then
                LBx = ContentT.FindControl("lb_qcperson")
                LBx.Items.Clear()
                LBx.Items.Add("")
                Do While (drsap.Read())
                    LBx.Items.Add(drsap(0))
                Loop
            End If
            drsap.Close()
            connsap.Close()
            LBx.SelectedIndex = 0
        End If

        SqlCmd = "SELECT T1.[Name] FROM OSCP T1 ORDER BY T1.Name"
        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (drsap.HasRows) Then
            LBx = ContentT.FindControl("lb_problemtype")
            LBx.Items.Clear()
            LBx.Items.Add("")
            Do While (drsap.Read())
                LBx.Items.Add(drsap(0))
            Loop
        End If
        drsap.Close()
        connsap.Close()
        LBx.SelectedIndex = 0

        SqlCmd = "SELECT T1.[Name] FROM OSCO T1 ORDER BY T1.Name"
        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (drsap.HasRows) Then
            LBx = ContentT.FindControl("lb_typedescrip")
            LBx.Items.Clear()
            LBx.Items.Add("")
            Do While (drsap.Read())
                LBx.Items.Add(drsap(0))
            Loop
        End If
        drsap.Close()
        connsap.Close()
        LBx.SelectedIndex = 0
    End Sub
    'Sub WriteListBoxItemForXFMRT()
    '    Dim LBx As ListBox
    '    Dim model, mdesc, mtype As String
    '    SqlCmd = "SELECT T1.[Name] FROM dbo.[OSCT] T1 ORDER BY T1.Name"
    '    drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
    '    If (drsap.HasRows) Then
    '        LBx = ContentT.FindControl("lb_machinetype")
    '        LBx.Items.Clear()
    '        LBx.Items.Add("")
    '        Do While (drsap.Read())
    '            LBx.Items.Add(drsap(0))
    '        Loop
    '    End If
    '    drsap.Close()
    '    connsap.Close()

    '    SqlCmd = "SELECT T0.u_model,T0.u_mdesc,T0.u_mtype " &
    '                "FROM dbo.[@UMMD] T0 order by T0.u_model,T0.u_mcode"
    '    drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
    '    LBx = ContentT.FindControl("lb_model")
    '    LBx.Items.Clear()
    '    LBx.Items.Add("")
    '    If (drsap.HasRows) Then
    '        Do While (drsap.Read())
    '            model = drsap(0)
    '            mdesc = drsap(1)
    '            mtype = drsap(2)
    '            LBx.Items.Add(mtype & "-" & model & "-" & mdesc)
    '        Loop
    '    End If
    '    LBx.SelectedIndex = 0
    '    drsap.Close()
    '    connsap.Close()

    '    SqlCmd = "SELECT distinct T0.cusname " &
    '            "FROM dbo.[@XFMRT] T0"
    '    drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
    '    If (drsap.HasRows) Then
    '        LBx = ContentT.FindControl("lb_cusname")
    '        LBx.Items.Clear()
    '        LBx.Items.Add("")
    '        Do While (drsap.Read())
    '            LBx.Items.Add(drsap(0))
    '        Loop
    '    End If
    '    drsap.Close()
    '    connsap.Close()
    '    LBx.SelectedIndex = 0

    '    SqlCmd = "SELECT distinct T0.qcperson " &
    '            "FROM dbo.[@XFMRT] T0"
    '    drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
    '    If (drsap.HasRows) Then
    '        LBx = ContentT.FindControl("lb_qcperson")
    '        LBx.Items.Clear()
    '        LBx.Items.Add("")
    '        Do While (drsap.Read())
    '            LBx.Items.Add(drsap(0))
    '        Loop
    '    End If
    '    drsap.Close()
    '    connsap.Close()
    '    LBx.SelectedIndex = 0

    '    SqlCmd = "SELECT T1.[Name] FROM OSCP T1 ORDER BY T1.Name"
    '    drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
    '    If (drsap.HasRows) Then
    '        LBx = ContentT.FindControl("lb_problemtype")
    '        LBx.Items.Clear()
    '        LBx.Items.Add("")
    '        Do While (drsap.Read())
    '            LBx.Items.Add(drsap(0))
    '        Loop
    '    End If
    '    drsap.Close()
    '    connsap.Close()
    '    LBx.SelectedIndex = 0

    '    SqlCmd = "SELECT T1.[Name] FROM OSCO T1 ORDER BY T1.Name"
    '    drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
    '    If (drsap.HasRows) Then
    '        LBx = ContentT.FindControl("lb_typedescrip")
    '        LBx.Items.Clear()
    '        LBx.Items.Add("")
    '        Do While (drsap.Read())
    '            LBx.Items.Add(drsap(0))
    '        Loop
    '    End If
    '    drsap.Close()
    '    connsap.Close()
    '    LBx.SelectedIndex = 0
    'End Sub
End Class