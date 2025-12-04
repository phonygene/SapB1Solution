Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms.OpenFileDialog
Imports AjaxControlToolkit
Imports System.IO
Public Class iqc
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public connsap As New SqlConnection
    Public connsaphead, connsapitem, connsaprecord As New SqlConnection
    Public SqlCmd As String
    Public oCompany As New SAPbobsCOM.Company
    Public dr, drsap, drsaphead, drsapitem, drsaprecord As SqlDataReader
    Public permsp201 As String
    Public ret As Long
    Public mode As String
    Public RBLMode As RadioButtonList
    Public BtnDraw, BtnUpload As Button
    Public BtnUpdate As Button
    Public BtnDelete As Button
    Public ponum_back, num_back As Long
    Public itemcode_back, itemname_back As String
    Public iqctype As Integer
    Public rest_inamount, po_amount As Integer
    Public stdwidth, stdheight As Integer
    Public ChkDelDW As CheckBox
    Public dwfile_exist As Boolean
    'Public preddlfun As Integer '因ddlfun=1時,無result textbox , 當把ddlfun切入=2時 , 因之前textbox為出來,故在後續填値時,會找不到而發生error,故用此flag來判斷
    Public permsqc100 As String
    Public FileUL As FileUpload
    Dim url As String
    'Public ScriptManager1 As New ScriptManager
    Public Structure CTItem_Data
        Dim row As Integer
        Dim icode As Long
    End Structure
    Public CTItem(20) As CTItem_Data

    'Public itemcode As String
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Dim attachflag As Boolean
        'attachflag = True
        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        dwfile_exist = False
        permsqc100 = CommUtil.GetAssignRight("qc100", Session("s_id"))
        stdwidth = 50
        stdheight = 20
        CTItem(0).row = 99
        If (Not IsPostBack) Then
            mode = Request.QueryString("mode")
            iqctype = Request.QueryString("iqctype")
            ViewState("mode") = mode
        End If
        FTCreate()
        If (IsPostBack) Then
            BtnDelete.OnClientClick = "return confirm('要刪除嗎')"
            BtnUpdate.OnClientClick = "return confirm('要新增或修改嗎')"

            iqctype = ViewState("iqctype")
            mode = ViewState("mode")
            Dim controlobj As Control
            controlobj = CommUtil.GetPostBackControl(Page) ' or CommUtil.GetPostBackControl(sender)
            If (controlobj IsNot Nothing) Then
                If (controlobj.ID = "rbl_mode") Then
                    If (mode = "edit") Then
                        mode = "showvalue"
                    Else
                        mode = "edit"
                    End If
                ElseIf (controlobj.ID = "btn_upload") Then
                    'attachflag = False
                End If
            End If

            If (iqctype <> 0) Then
                If (iqctype = 1) Then 'MIT(主檔) 顯示
                    itemcode_back = ViewState("itemcode_back")
                    itemname_back = ViewState("itemname_back")
                ElseIf (iqctype = 2) Then 'MIT(主檔) 建立
                    'MsgBox("post " & ViewState("itemname_back"))
                    itemcode_back = ViewState("itemcode_back")
                    itemname_back = ViewState("itemname_back")
                ElseIf (iqctype = 3) Then '新建IQT
                    ponum_back = ViewState("ponum_back")
                    itemcode_back = ViewState("itemcode_back")
                    itemname_back = ViewState("itemname_back")
                    rest_inamount = ViewState("rest_inamount")
                    po_amount = ViewState("po_amount")
                ElseIf (iqctype = 4) Then '顯示IQT
                    num_back = ViewState("num_back")
                    itemname_back = ViewState("itemname_back")
                End If
            Else
                'nothing
            End If
        Else
            If (iqctype = 3) Then
                itemcode_back = Request.QueryString("itemcode")
                itemname_back = Request.QueryString("itemname")
                ponum_back = Request.QueryString("po")
                rest_inamount = Request.QueryString("rest_inamount")
                po_amount = Request.QueryString("po_amount")
            ElseIf (iqctype = 4) Then
                num_back = Request.QueryString("docnum")
                itemname_back = Request.QueryString("itemname")
            ElseIf (iqctype = 1 Or iqctype = 2) Then
                itemcode_back = Request.QueryString("itemcode")
                itemname_back = Request.QueryString("itemname")
            End If
            If (Request.QueryString("act") = "fileupload") Then
                CommUtil.ShowMsg(Me, "圖檔上傳成功")
            ElseIf (Request.QueryString("act") = "filedelete") Then
                CommUtil.ShowMsg(Me, "圖檔刪除成功")
            End If
        End If
        CTCreate()
        If (Not IsPostBack) Then
            If (mode <> "edit") Then
                RBLMode.SelectedValue = 1
            Else
                RBLMode.SelectedValue = 2
                BtnUpdate.Enabled = True
                If (iqctype = 1 Or iqctype = 4) Then
                    BtnDelete.Enabled = True
                End If
            End If
            GetActionPara1()
            DisplayIQCData()
        End If
        '某目錄取檔sample
        'Dim strPath As String = "C:\QC\圖檔\"
        'Dim di As DirectoryInfo
        'di = New DirectoryInfo(strPath)
        'Dim fi As FileInfo() = di.GetFiles("14.*", SearchOption.AllDirectories) ' 含子目錄
        'Dim fi As FileInfo() = di.GetFiles(CT.Rows(0).Cells(3).Text & ".*")
        'MsgBox(fi(0).Name)
        url = Application("http")
        Dim targetDir, displayname As String
        Dim targetFile As String
        targetDir = HttpContext.Current.Server.MapPath("~/") & "\AttachFile\QC\DW\"
        Dim di As DirectoryInfo
        di = New DirectoryInfo(targetDir)
        Dim fi As FileInfo() = di.GetFiles(CT.Rows(0).Cells(3).Text & ".*")
        If (fi.Length = 0) Then
            displayname = ""
        Else
            displayname = fi(0).Name
        End If
        targetFile = targetDir & displayname
        'If (File.Exists(targetFile) And attachflag = True) Then
        If (File.Exists(targetFile)) Then
            Dim httpfile As String
            httpfile = url & "AttachFile/QC/DW/" & fi(0).Name
            iframeContent.Attributes.Remove("src")
            iframeContent.Attributes.Add("src", httpfile)
            dwfile_exist = True
        Else
            iframeContent.Visible = False
        End If
    End Sub

    Sub FTCreate()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim Labelx As Label
        Dim Hyper As HyperLink
        Dim funindex, indexpage As Integer

        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = New TableCell()
        'tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center

        funindex = Request.QueryString("funindex")
        indexpage = Request.QueryString("indexpage")
        Hyper = New HyperLink()
        Hyper.Text = "回總表"
        If (funindex = 1) Then
            Hyper.NavigateUrl = "qc.aspx?smid=qc&smode=0&funindex=" & funindex & "&kw=" & Request.QueryString("kw") & "&indexpage=" & indexpage
        ElseIf (funindex = 3) Then
            Hyper.NavigateUrl = "qc.aspx?smid=qc&smode=0&funindex=" & funindex & "&po=" & Request.QueryString("po") & "&indexpage=" & indexpage
        Else
            Hyper.NavigateUrl = "qc.aspx?smid=qc&smode=0&funindex=" & funindex & "&indexpage=" & indexpage
        End If

        Hyper.Font.Underline = False
        Hyper.ID = "hyper_backiqclist"
        tCell.Controls.Add(Hyper)
        tRow.Cells.Add(tCell)

        tCell = New TableCell()
        'tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        RBLMode = New RadioButtonList()
        RBLMode.ID = "rbl_mode"
        RBLMode.Items.Add("顯示")
        RBLMode.Items.Add("編輯")
        RBLMode.Items(0).Value = 1
        RBLMode.Items(1).Value = 2
        RBLMode.Font.Size = 10
        'RBLMode.Width = 100
        RBLMode.RepeatDirection = RepeatDirection.Vertical
        RBLMode.AutoPostBack = True
        CommUtil.DisableObjectByPermission(RBLMode, permsqc100, "m")
        AddHandler RBLMode.SelectedIndexChanged, AddressOf RBLMode_SelectedIndexChanged
        tCell.Controls.Add(RBLMode)

        tRow.Cells.Add(tCell)
        '--------------------------------------------------------------------
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        Labelx = New Label()
        Labelx.ID = "label_adds4"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        BtnUpdate = New Button()
        BtnUpdate.ID = "btn_upd"
        'CommUtil.DisableObjectByPermission(Btnx, permsmf204, "n")
        BtnUpdate.Text = "更新"
        BtnUpdate.Enabled = False
        AddHandler BtnUpdate.Click, AddressOf BtnUpdate_Click
        tCell.Controls.Add(BtnUpdate)

        Labelx = New Label()
        Labelx.ID = "label_adds5"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        BtnDelete = New Button()
        BtnDelete.ID = "btn_del"
        'CommUtil.DisableObjectByPermission(Btnx, permsmf204, "n")
        BtnDelete.Text = "刪除"
        BtnDelete.Enabled = False
        AddHandler BtnDelete.Click, AddressOf BtnDelete_Click
        tCell.Controls.Add(BtnDelete)

        tRow.Cells.Add(tCell)
        FT.Rows.Add(tRow)

        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = New TableCell()
        tCell.ColumnSpan = 3
        tCell.HorizontalAlign = HorizontalAlign.Left

        Labelx = New Label()
        Labelx.ID = "label_fileul"
        Labelx.Text = "選擇上傳之圖檔(pdf)檔案"
        tCell.Controls.Add(Labelx)
        FileUL = New FileUpload()
        FileUL.ID = "fileul"
        tCell.Controls.Add(FileUL)

        ChkDelDW = New CheckBox
        ChkDelDW.ID = "chk_del_0"
        ChkDelDW.Text = "刪除圖檔"
        ChkDelDW.AutoPostBack = True
        AddHandler ChkDelDW.CheckedChanged, AddressOf ChkDelDW_CheckedChanged
        tCell.Controls.Add(ChkDelDW)

        Labelx = New Label()
        Labelx.ID = "label_upfile0"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)

        BtnUpload = New Button()
        BtnUpload.ID = "btn_upload"
        CommUtil.DisableObjectByPermission(BtnUpload, permsqc100, "m")
        BtnUpload.Text = "上傳圖檔"
        BtnUpload.OnClientClick = "return confirm('若有編輯需先儲存,否則資料會Loss,要繼續嗎?')"
        AddHandler BtnUpload.Click, AddressOf BtnUpload_Click
        tCell.Controls.Add(BtnUpload)

        'ChkDel = New CheckBox
        'ChkDel.ID = "chk_del"
        'ChkDel.Text = "刪除聯絡單"
        'ChkDel.AutoPostBack = True
        'AddHandler ChkDel.CheckedChanged, AddressOf ChkDel_CheckedChanged
        'tCell.Controls.Add(ChkDel)
        tRow.Cells.Add(tCell)

        FT.Rows.Add(tRow)
    End Sub
    Protected Sub ChkDelDW_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)
        If (ChkDelDW.Checked) Then
            If dwfile_exist Then
                BtnUpload.Text = "刪除圖檔"
            Else
                ChkDelDW.Checked = False
                CommUtil.ShowMsg(Me, "無圖檔,無法執行刪除")
            End If
            BtnUpload.OnClientClick = ""
        Else
            BtnUpload.Text = "上傳圖檔"
        End If
    End Sub
    Sub CTCreate()
        CTHead1()
        CTHead2()
        CTHead3()
        CTRecordHead()
        CTRecord()
        CTTail_1()
        CTTail_2()
        CTTail_3()
        CTTail_4()
        CTTail_5()
        CreateRowOfFullCell()
    End Sub

    Function CellSet(text As String, rowspan As Integer, colspan As Integer, FondBold As Boolean, txtid As String, width As Integer, height As Integer)
        Dim tCell As TableCell
        tCell = New TableCell()
        tCell.BorderWidth = 1
        If (txtid = "amemo" Or txtid = "cmemo") Then
            tCell.HorizontalAlign = HorizontalAlign.Left
        Else
            tCell.HorizontalAlign = HorizontalAlign.Center
        End If

        tCell.Wrap = True
        If (mode = "showvalue" Or txtid = "") Then
            tCell.Text = text
        End If
        tCell.ColumnSpan = colspan
        tCell.RowSpan = rowspan
        tCell.Width = width 'stdwidth * colspan * 0.95
        tCell.Height = height '20 * rowspan
        tCell.Font.Bold = FondBold
        If (mode = "edit" And txtid <> "") Then
            tCell.Controls.Add(TextBoxSet(txtid, text, width, height))
        End If
        Return tCell
    End Function

    Function TextBoxSet(id As String, text As String, width As Integer, height As Integer)
        Dim Txtx As TextBox
        Dim str() As String
        str = Split(id, "_")
        Txtx = New TextBox()
        Txtx.ID = "txt_" & id
        Txtx.Text = text
        Txtx.Width = width * 0.95
        Txtx.Height = height
        If (id = "amemo" Or id = "cmemo") Then
            Txtx.TextMode = TextBoxMode.MultiLine
            Txtx.Width = width * 0.95 + 120
        End If
        If (str(0) = "iresult") Then
            Txtx.AutoPostBack = True
            If (iqctype = 1 Or iqctype = 2) Then
                Txtx.Enabled = False
            Else
                Txtx.Enabled = True
            End If
            AddHandler Txtx.TextChanged, AddressOf TxtResult_TextChanged
        End If
        'If (id = "inspecdate") Then
        '    If (drsaphead(12) = 0) Then
        '        Txtx.Enabled = False
        '    Else
        '        Txtx.Enabled = True
        '    End If
        'End If
        'If (id = "auditdate") Then
        '    If (drsaphead(16) = "1900/1/1") Then
        '        Txtx.Enabled = False
        '    Else
        '        Txtx.Enabled = True
        '    End If
        'End If
        Return Txtx
    End Function

    Sub CreateRowOfFullCell()
        Dim i As Integer
        Dim tCell As TableCell
        Dim tRow As TableRow

        tRow = New TableRow()
        tRow.BorderWidth = 0
        For i = 1 To 18
            tCell = CellSet("", 1, 1, False, "", stdwidth, stdheight)
            tCell.BorderWidth = 0
            tCell.Height = 1
            tRow.Cells.Add(tCell)
        Next
        CT.Rows.Add(tRow)
    End Sub
    Sub CTTail_1()
        Dim tCell As TableCell
        Dim tRow As TableRow
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Left
        tCell.Wrap = True
        tCell.Font.Bold = True
        tCell.ColumnSpan = 18
        tCell.Text = "檢驗儀器 : C(分釐卡尺)&nbsp&nbsp&nbsp&nbspV(2.5D量測儀)&nbsp&nbsp&nbsp&nbspH(高度規)" &
                     "&nbsp&nbsp&nbsp&nbspP(塞規)&nbsp&nbsp&nbsp&nbspM(外徑測微器)&nbsp&nbsp&nbsp&nbspI(內徑測微器)" &
                     "&nbsp&nbsp&nbsp&nbspR(卡尺)&nbsp&nbsp&nbsp&nbspX(手持便攜式3D量測儀)&nbsp&nbsp&nbsp&nbspO(其它)"
        tRow.Cells.Add(tCell)
        CT.Rows.Add(tRow)
    End Sub
    Sub CTTail_2()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim rRBL As RadioButtonList
        Dim judge As Integer
        judge = 0
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Cells.Add(CellSet("綜合判定", 1, 2, True, "", stdwidth * 2, stdheight * 1))
        tCell = CellSet("", 1, 17, False, "", stdwidth * 17, stdheight * 1)
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_judge"
        rRBL.Items.Add("允收")
        rRBL.Items.Add("拒收")
        rRBL.Items.Add("特採")
        rRBL.Items.Add("其他")
        rRBL.Items(0).Value = 1
        rRBL.Items(1).Value = 2
        rRBL.Items(2).Value = 3
        rRBL.Items(3).Value = 4
        rRBL.RepeatDirection = RepeatDirection.Vertical
        'rRBL.SelectedValue = judge
        If (iqctype = 3) Then
            rRBL.SelectedIndex = -1
            rRBL.Enabled = True
        Else
            rRBL.SelectedIndex = 0
            rRBL.Enabled = False
        End If
        tCell.Controls.Add(rRBL)
        tRow.Cells.Add(tCell)
        CT.Rows.Add(tRow)
    End Sub
    Sub CTTail_3()
        Dim tRow As TableRow
        Dim tCell As TableCell
        Dim BtnInspector, BtnAuditor As Button
        Dim inspector, auditor, cmemo As String
        inspector = "簽名"
        auditor = "簽名"
        cmemo = ""
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Cells.Add(CellSet("檢查<br>記事", 2, 1, True, "", stdwidth * 1, stdheight * 2))
        tRow.Cells.Add(CellSet(cmemo, 2, 13, False, "cmemo", stdwidth * 13, stdheight * 2 + 10))
        tRow.Cells.Add(CellSet("檢查人員", 1, 2, True, "", stdwidth * 2, stdheight * 1))
        tRow.Cells.Add(CellSet("審查人員", 1, 2, True, "", stdwidth * 2, stdheight * 1))
        CT.Rows.Add(tRow)
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = CellSet("", 1, 2, False, "", stdwidth * 2, stdheight * 1 + 10)
        If (mode = "edit") Then
            BtnInspector = New Button
            BtnInspector.ID = "btn_inspector"
            BtnInspector.Text = inspector
            BtnInspector.Width = stdwidth * 2
            BtnInspector.Height = stdheight + 10
            CommUtil.DisableObjectByPermission(BtnInspector, permsqc100, "m")
            tCell.Controls.Add(BtnInspector)
            AddHandler BtnInspector.Click, AddressOf BtnInspector_Click
        End If
        tRow.Cells.Add(tCell)
        'tRow.Cells.Add(CellSet(inspector, 1, 2, False, "", stdwidth * 2, stdheight * 1))
        'tRow.Cells.Add(CellSet(auditor, 1, 2, False, "", stdwidth * 2, stdheight * 1))
        tCell = CellSet("", 1, 2, False, "", stdwidth * 2, stdheight * 1 + 10)
        If (mode = "edit") Then
            BtnAuditor = New Button
            BtnAuditor.ID = "btn_auditor"
            BtnAuditor.Text = auditor
            BtnAuditor.Width = stdwidth * 2
            BtnAuditor.Height = stdheight * 1 + 10
            CommUtil.DisableObjectByPermission(BtnAuditor, permsqc100, "a")
            tCell.Controls.Add(BtnAuditor)
            AddHandler BtnAuditor.Click, AddressOf BtnAuditor_Click
        End If
        tRow.Cells.Add(tCell)
        CT.Rows.Add(tRow)
    End Sub
    Sub CTTail_4()
        Dim tRow As TableRow
        Dim inspecdate, auditdate, amemo As String
        inspecdate = ""
        auditdate = ""
        amemo = ""
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Cells.Add(CellSet("審核<br>記事", 2, 1, True, "", stdwidth * 1, stdheight * 2))
        tRow.Cells.Add(CellSet(amemo, 2, 13, False, "amemo", stdwidth * 13, stdheight * 2 + 10))
        tRow.Cells.Add(CellSet("日期", 1, 2, True, "", stdwidth * 2, stdheight * 1))
        tRow.Cells.Add(CellSet("日期", 1, 2, True, "", stdwidth * 2, stdheight * 1))
        CT.Rows.Add(tRow)
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Cells.Add(CellSet(inspecdate, 1, 2, False, "inspecdate", stdwidth * 2, stdheight * 1))
        tRow.Cells.Add(CellSet(auditdate, 1, 2, False, "auditdate", stdwidth * 2, stdheight * 1))
        CT.Rows.Add(tRow)
    End Sub
    Sub CTTail_5()
        Dim tCell As TableCell
        Dim tRow As TableRow
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Left
        tCell.Wrap = True
        tCell.ColumnSpan = 18
        tCell.Font.Bold = True
        tCell.Text = "說明:<br>不可量化檢驗項目以 OK , NG 代表 , 可量化檢驗項目以量測值填寫<br>" &
                     "驗收標準及方法依作業標準 , 不合格應被住處理方式"
        tRow.Cells.Add(tCell)
        CT.Rows.Add(tRow)
    End Sub
    Sub CTHead1()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim rRBL As RadioButtonList
        Dim mapno, docnum As String
        Dim inamount, famount As String
        Dim firstqc As Integer
        firstqc = 1
        inamount = ""
        famount = ""
        mapno = ""
        docnum = ""
        'If (mode <> "showempty") Then
        '    If (iqctype = 1) Then 'MIT 顯示
        '        firstqc = 0
        '        inamount = ""
        '        famount = ""
        '        mapno = CStr(drsaphead(1))
        '        docnum = ""
        '    ElseIf (iqctype = 2) Then 'MIT 建立
        '        firstqc = 0
        '        inamount = ""
        '        famount = ""
        '        mapno = ""
        '        docnum = ""
        '    ElseIf (iqctype = 3) Then 'IQT新建立
        '        firstqc = 0
        '        inamount = ""
        '        famount = ""
        '        mapno = CStr(drsaphead(1))
        '        docnum = ""
        '    ElseIf (iqctype = 4) Then 'IQT 顯示
        '        firstqc = CStr(drsaphead(4))
        '        inamount = CStr(drsaphead(0))
        '        famount = CStr(drsaphead(1))
        '        mapno = CStr(drsaphead(19))
        '        docnum = CStr(drsaphead(2))
        '    ElseIf (iqctype = 5) Then 'IQT續建立
        '        firstqc = 0
        '        inamount = rest_inamount
        '        famount = ""
        '        mapno = CStr(drsaphead(1))
        '        docnum = ""

        '    End If
        'End If

        tRow = New TableRow()
        tRow.BorderWidth = 1

        tRow.Cells.Add(CellSet("表單編號", 1, 2, True, "", stdwidth * 2, stdheight * 1))
        tRow.Cells.Add(CellSet(docnum, 1, 2, False, "", stdwidth * 2, stdheight * 1))

        tRow.Cells.Add(CellSet("對應數字", 1, 2, True, "", stdwidth * 2, stdheight * 1))

        tCell = CellSet("", 1, 2, False, "", stdwidth * 2, stdheight * 1)
        'BtnDraw = New Button()
        'BtnDraw.ID = "btn_draw"
        'CommUtil.DisableObjectByPermission(BtnDraw, permsqc100, "nm")
        'BtnDraw.Width = stdwidth * 2
        'AddHandler BtnDraw.Click, AddressOf BtnDraw_Click
        'tCell.Controls.Add(BtnDraw)
        tRow.Cells.Add(tCell)

        tRow.Cells.Add(CellSet("此次進量", 1, 2, True, "", stdwidth * 2, stdheight * 1))
        tRow.Cells.Add(CellSet(inamount, 1, 1, False, "inamount", stdwidth * 1, stdheight * 1))

        tRow.Cells.Add(CellSet("已檢數量", 1, 2, True, "", stdwidth * 2, stdheight * 1))
        tRow.Cells.Add(CellSet(famount, 1, 1, False, "famount", stdwidth * 1, stdheight * 1))

        tCell = CellSet("", 1, 4, True, "", stdwidth * 4, stdheight * 1)
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_firstqc"
        rRBL.Items.Add("初驗")
        rRBL.Items.Add("復驗")
        rRBL.Items(0).Value = 1
        rRBL.Items(1).Value = 2
        rRBL.RepeatDirection = RepeatDirection.Vertical
        If (iqctype = 3) Then
            rRBL.SelectedIndex = -1
            rRBL.Enabled = True
        Else
            rRBL.SelectedIndex = 0
            rRBL.Enabled = False
        End If
        tCell.Controls.Add(rRBL)
        tRow.Cells.Add(tCell)

        CT.Rows.Add(tRow)
    End Sub

    Sub CTHead2()
        'Dim SqlCmd As String
        'Dim drinner As SqlDataReader
        Dim tRow As TableRow
        Dim itemcode, itemname, vender, createdate As String
        itemcode = ""
        itemname = ""
        vender = ""
        createdate = ""
        'If (mode <> "showempty") Then
        '    If (iqctype = 1 Or iqctype = 2) Then 'MIT 顯示 , 建立
        '        itemcode = itemcode_back
        '        itemname = drsaphead(0)
        '        vender = ""
        '        createdate = ""
        '    ElseIf (iqctype = 3 Or iqctype = 5) Then 'IQT建立
        '        itemcode = itemcode_back
        '        itemname = itemname_back
        '        vender = drsaphead(1)
        '        createdate = Format(Now(), "yyyy/MM/dd")
        '    ElseIf (iqctype = 4) Then 'IQT 顯示
        '        itemcode = drsaphead(3)
        '        'SqlCmd = "SELECT T0.itemname FROM OITM T0 where T0.itemcode='" & itemcode & "'"
        '        'drinner = CommUtil.SelectSqlUsingDr(SqlCmd, connsaphead)
        '        'drinner.Read()
        '        'itemname = drinner(0)
        '        itemname = itemname_back
        '        vender = drsaphead(5)
        '        createdate = drsaphead(6)
        '        'drinner.Close()
        '    End If
        'End If
        tRow = New TableRow()
        tRow.BorderWidth = 1

        tRow.Cells.Add(CellSet("料號", 1, 1, True, "", stdwidth * 1, stdheight * 1))
        tRow.Cells.Add(CellSet(itemcode, 1, 3, False, "", stdwidth * 3, stdheight * 1))

        tRow.Cells.Add(CellSet("品名", 1, 1, True, "", stdwidth * 1, stdheight * 1))
        tRow.Cells.Add(CellSet(itemname, 1, 6, False, "", stdwidth * 6, stdheight * 1))

        tRow.Cells.Add(CellSet("廠商", 1, 1, True, "", stdwidth * 1, stdheight * 1))
        tRow.Cells.Add(CellSet(vender, 1, 2, False, "vender", stdwidth * 2, stdheight * 1))

        tRow.Cells.Add(CellSet("建立日期", 1, 2, True, "", stdwidth * 2, stdheight * 1))
        tRow.Cells.Add(CellSet(createdate, 1, 2, False, "", stdwidth * 2, stdheight * 1))

        CT.Rows.Add(tRow)
    End Sub

    Sub CTHead3()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim rRBL As RadioButtonList
        Dim po As String
        Dim poamount, tamount, ngamount As String
        Dim mtype As Integer
        tamount = ""
        ngamount = ""
        mtype = 0
        po = ""
        poamount = ""
        'If (mode <> "showempty") Then
        '    If (iqctype = 1 Or iqctype = 2) Then 'MIT 顯示 , 建立
        '        'nothing
        '    ElseIf (iqctype = 3) Then 'IQT建立
        '        tamount = CStr(drsaphead(0))
        '        po = ponum_back
        '        poamount = CStr(drsaphead(0))
        '    ElseIf (iqctype = 4) Then 'IQT 顯示
        '        tamount = CStr(drsaphead(9))
        '        ngamount = CStr(drsaphead(10))
        '        mtype = drsaphead(11)
        '        po = CStr(drsaphead(8))
        '        poamount = CStr(drsaphead(7))
        '    ElseIf (iqctype = 5) Then 'IQT建立
        '        tamount = rest_inamount
        '        po = ponum_back
        '        poamount = CStr(drsaphead(0))
        '    End If
        'End If

        tRow = New TableRow()
        tRow.BorderWidth = 1

        tRow.Cells.Add(CellSet("PO", 1, 1, True, "", stdwidth * 1, stdheight * 1))
        tRow.Cells.Add(CellSet(po, 1, 2, False, "ponum", stdwidth * 2, stdheight * 1))

        tRow.Cells.Add(CellSet("採購數量", 1, 2, True, "", stdwidth * 2, stdheight * 1))
        tRow.Cells.Add(CellSet(poamount, 1, 1, False, "poamount", stdwidth * 1, stdheight * 1))

        tRow.Cells.Add(CellSet("抽驗數", 1, 1, True, "", stdwidth * 1, stdheight * 1))
        tRow.Cells.Add(CellSet(tamount, 1, 1, False, "tamount", stdwidth * 1, stdheight * 1))

        tRow.Cells.Add(CellSet("不良數", 1, 1, True, "", stdwidth * 1, stdheight * 1))
        tRow.Cells.Add(CellSet(ngamount, 1, 1, False, "ngamount", stdwidth * 1, stdheight * 1))

        tRow.Cells.Add(CellSet("進料類別", 1, 2, True, "", stdwidth * 2, stdheight * 1))
        tCell = CellSet("", 1, 6, False, "", stdwidth * 6, stdheight * 1)
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_mtype"
        rRBL.Items.Add("銑件")
        rRBL.Items.Add("車件")
        rRBL.Items.Add("市購")
        rRBL.Items.Add("鈑金")
        rRBL.Items.Add("骨架")
        rRBL.Items(0).Value = 1
        rRBL.Items(1).Value = 2
        rRBL.Items(2).Value = 3
        rRBL.Items(3).Value = 4
        rRBL.Items(4).Value = 5
        rRBL.RepeatDirection = RepeatDirection.Vertical
        If (iqctype = 3) Then
            rRBL.SelectedIndex = -1
            rRBL.Enabled = True
        Else
            rRBL.SelectedIndex = 0
            rRBL.Enabled = False
        End If
        tCell.Controls.Add(rRBL)
        tRow.Cells.Add(tCell)
        CT.Rows.Add(tRow)
    End Sub

    Sub CTRecordHead()
        'Dim tCell As TableCell
        Dim tRow As TableRow
        'Dim stdwidth As Integer
        'stdwidth = 50
        Dim i As Integer

        tRow = New TableRow()
        tRow.BorderWidth = 1

        tRow.Cells.Add(CellSet("項次", 2, 1, True, "", stdwidth * 2, stdheight * 1))
        tRow.Cells.Add(CellSet("檢驗項目", 2, 3, True, "", stdwidth * 3, stdheight * 2))
        tRow.Cells.Add(CellSet("標準", 2, 2, True, "", stdwidth * 2, stdheight * 2))
        tRow.Cells.Add(CellSet("誤差", 2, 2, True, "", stdwidth * 2, stdheight * 2))
        tRow.Cells.Add(CellSet("檢驗結果(量測值記錄)", 1, 9, True, "", stdwidth * 9, stdheight * 1))
        tRow.Cells.Add(CellSet("使用工具", 2, 1, True, "", stdwidth * 2, stdheight * 1))
        CT.Rows.Add(tRow)
        tRow = New TableRow()
        tRow.BorderWidth = 1
        For i = 1 To 9
            tRow.Cells.Add(CellSet(CStr(i), 1, 1, True, "", stdwidth * 1, stdheight * 1))
        Next
        CT.Rows.Add(tRow)
    End Sub
    Sub CTRecord()
        Dim tRow As TableRow
        Dim iname, ispec, itol, iresult, itool As String
        Dim i, j As Integer
        Dim iid_num, rid_num As String
        For i = 1 To 14
            iname = ""
            ispec = ""
            itol = ""
            itool = ""
            iid_num = CStr(i)
            tRow = New TableRow()
            tRow.BorderWidth = 1
            tRow.Cells.Add(CellSet(CStr(i), 1, 1, True, "", stdwidth * 1, stdheight * 1))

            tRow.Cells.Add(CellSet(iname, 1, 3, True, "iname_" & iid_num, stdwidth * 3, stdheight * 1))

            tRow.Cells.Add(CellSet(ispec, 1, 2, True, "ispec_" & iid_num, stdwidth * 2, stdheight * 1))

            tRow.Cells.Add(CellSet(itol, 1, 2, True, "itol_" & iid_num, stdwidth * 2, stdheight * 1))
            For j = 1 To 9
                rid_num = CStr(i) & "_" & CStr(j)
                iresult = ""
                'If (DDLFun.SelectedIndex = 0 Or DDLFun.SelectedIndex = 1) Then
                'tRow.Cells.Add(CellSet(iresult, 1, 1, False, "", stdwidth * 1, stdheight * 1))
                'Else
                'tRow.Cells.Add(CellSet(iresult, 1, 1, False, "iresult_" & rid_num, stdwidth * 1, stdheight * 1))
                'End If
                tRow.Cells.Add(CellSet(iresult, 1, 1, False, "iresult_" & rid_num, stdwidth * 1, stdheight * 1))
            Next
            tRow.Cells.Add(CellSet(itool, 1, 1, True, "itool_" & iid_num, stdwidth * 1, stdheight * 1))
            CT.Rows.Add(tRow)
        Next
    End Sub


    'Sub GetActionPara()
    '    If (DDLRecord.SelectedIndex = 0) Then
    '        ViewState("iqctype") = 0
    '        iqctype = 0
    '        'ViewState("mode") = "showempty"
    '        Exit Sub
    '    End If
    '    Dim str() As String
    '    Dim str1() As String
    '    Dim status_str, numstr As String
    '    Dim po As Long
    '    Dim tinamount, poamount As Integer
    '    str = Split(DDLRecord.SelectedValue, "@")
    '    If (DDLFun.SelectedIndex = 1) Then
    '        status_str = str(0)
    '        If (status_str = "已建立") Then '顯示MIT
    '            ViewState("itemcode_back") = str(1)
    '            ViewState("itemname_back") = str(2)
    '            ViewState("iqctype") = 1
    '            itemcode_back = str(1)
    '            iqctype = 1
    '        Else '建立MIT
    '            ViewState("iqctype") = 2 '建立MIT
    '            ViewState("itemcode_back") = str(1)
    '            ViewState("itemname_back") = str(2)
    '            iqctype = 2
    '            itemcode_back = str(1)
    '            itemname_back = str(2)
    '        End If
    '    ElseIf (DDLFun.SelectedIndex = 2) Then
    '        status_str = str(0)
    '        numstr = str(1)
    '        str1 = Split(numstr, ":")
    '        If (status_str = "可建單") Then '以po 號 show id ,idname ,vender 來新增表單
    '            ViewState("iqctype") = 3 '新建IQT(未曾建過)
    '            ViewState("ponum_back") = str1(1)
    '            ViewState("itemcode_back") = str(2)
    '            ViewState("itemname_back") = str(3)
    '            iqctype = 3
    '            ponum_back = str1(1)
    '            itemcode_back = str(2)
    '            itemname_back = str(3)
    '        ElseIf (status_str = "需建基本資料") Then
    '            ViewState("iqctype") = 2 '建立MIT
    '            ViewState("itemcode_back") = str(2)
    '            ViewState("itemname_back") = str(3)
    '            DDLFun.SelectedIndex = 1
    '            iqctype = 2
    '            itemcode_back = str(2)
    '            itemname_back = str(3)
    '            DDLRecord.Items.Clear()
    '            DDLRecord.Items.Add("已篩選出單據 請選擇")
    '            DDLRecord.Items.Add("未建單@" & itemcode_back & "@" & itemname_back)
    '            DDLRecord.SelectedIndex = 1
    '        ElseIf (status_str = "檢驗中" Or status_str = "未審單" Or status_str = "未結案" Or status_str = "已結案") Then
    '            SqlCmd = "SELECT T0.u_po,T0.u_amount FROM dbo.[@UIQT] T0 " &
    '                    "where T0.u_docnum=" & str1(1)
    '            drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
    '            drsap.Read()
    '            po = drsap(0)
    '            poamount = drsap(1)
    '            drsap.Close()
    '            connsap.Close()
    '            SqlCmd = "SELECT sum(u_inamount) FROM dbo.[@UIQT] T0 " &
    '                    "where T0.u_itemcode='" & str(2) & "' and T0.u_po=" & po
    '            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
    '            dr.Read()
    '            tinamount = dr(0)
    '            dr.Close()
    '            connsap.Close()
    '            If (tinamount < poamount) Then
    '                If (cflag) Then '新建IQT(之前已建)
    '                    SqlCmd = "SELECT T0.u_tiname,T0.u_tispec,T0.u_tol,T0.u_tooluse,T0.Code " &
    '                            "FROM dbo.[@UMIT] T0 " &
    '                            "where T0.u_itemcode='" & str(2) & "' order by T0.u_tiseq"
    '                    drsapitem = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsapitem)
    '                    If (Not (drsapitem.HasRows)) Then
    '                        CommUtil.ShowMsg(Me, str(2) & "還未建IQC主檔, 應是刪除了,請重建後再執行剩餘新增")
    '                        Exit Sub
    '                    End If
    '                    ViewState("iqctype") = 5
    '                    ViewState("ponum_back") = po
    '                    ViewState("rest_inamount") = poamount - tinamount
    '                    ViewState("itemcode_back") = str(2)
    '                    ViewState("itemname_back") = str(3)
    '                    iqctype = 5
    '                    ponum_back = po
    '                    rest_inamount = poamount - tinamount
    '                    itemcode_back = str(2)
    '                    itemname_back = str(3)
    '                    CommUtil.ShowMsg(Me, "接下執行建立IQC表單,數量剩" & poamount - tinamount & "個可建")
    '                    drsapitem.Close()
    '                    connsapitem.Close()
    '                Else '顯示IQT(4)
    '                    ViewState("num_back") = str1(1)
    '                    ViewState("itemname_back") = str(3)
    '                    ViewState("iqctype") = 4

    '                    num_back = str1(1)
    '                    itemname_back = str(3)
    '                    iqctype = 4
    '                    'MsgBox(str1(1) & "-" & str(3) & "-" & ViewState("iqctype"))
    '                End If
    '            Else '顯示IQT(4)
    '                ViewState("num_back") = str1(1)
    '                ViewState("itemname_back") = str(3)
    '                ViewState("iqctype") = 4

    '                num_back = str1(1)
    '                itemname_back = str(3)
    '                iqctype = 4
    '            End If
    '        End If
    '    Else

    '    End If
    '    If (RBLMode.SelectedValue = 1) Then
    '        ViewState("mode") = "showvalue"
    '    Else
    '        ViewState("mode") = "edit"
    '    End If

    'End Sub
    Sub GetActionPara1()
        If (iqctype = 1) Then
            '顯示MIT
            ViewState("itemcode_back") = itemcode_back
            ViewState("itemname_back") = itemname_back
            ViewState("iqctype") = 1
            'MsgBox("not post " & ViewState("itemname_back"))
        ElseIf (iqctype = 2) Then '建立MIT
            ViewState("iqctype") = 2 '建立MIT
            ViewState("itemcode_back") = itemcode_back
            ViewState("itemname_back") = itemname_back
        ElseIf (iqctype = 3) Then
            ViewState("iqctype") = 3
            ViewState("ponum_back") = ponum_back
            ViewState("itemcode_back") = itemcode_back
            ViewState("itemname_back") = itemname_back
            ViewState("rest_inamount") = rest_inamount
            ViewState("po_amount") = po_amount
        ElseIf (iqctype = 4) Then 'ooooo
            ViewState("num_back") = num_back
            ViewState("itemname_back") = itemname_back
            ViewState("iqctype") = 4
        End If
        If (RBLMode.SelectedValue = 1) Then
            ViewState("mode") = "showvalue"
        Else
            ViewState("mode") = "edit"
        End If
    End Sub
    Protected Sub RBLMode_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)
        'GetActionPara()
        If (RBLMode.SelectedValue = 1) Then
            ViewState("mode") = "showvalue"
            mode = "showvalue"
            If (RBLMode.SelectedIndex = 0) Then
                BtnUpdate.Enabled = False
            End If
        Else
            ViewState("mode") = "edit"
            mode = "edit"
            If (RBLMode.SelectedIndex <> 0) Then
                BtnUpdate.Enabled = True
                BtnDelete.Enabled = True
            End If
        End If
        DisplayIQCData()
    End Sub

    'Protected Sub BtnSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    '    If (TxtKW.Text <> "") Then
    '        If (DDLFun.SelectedIndex = 1) Then
    '            UMITKWSearch(TxtKW.Text)
    '        ElseIf (DDLFun.SelectedIndex = 2) Then
    '            UIQTKWSearch(TxtKW.Text)
    '        End If
    '    Else
    '        CommUtil.ShowMsg(Me, "需輸入關鍵字")
    '    End If
    'End Sub

    'Sub UIQTKWSearch(kw As String)
    '    Dim str As String
    '    SqlCmd = "SELECT T0.itemcode,T0.itemname,T0.u_F7 " &
    '            "FROM OITM T0 " &
    '            "where T0.itemcode like '%" & kw & "%' or T0.itemname like '%" & kw & "%' order by T0.itemcode"
    '    drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
    '    DDLRecord.Items.Clear()
    '    If (drsap.HasRows) Then
    '        DDLRecord.Items.Add("已篩選出單據 請選擇")
    '        Do While (drsap.Read())
    '            If (drsap(2) = 0) Then
    '                str = "需建基本資料@採購單:0@" & drsap(0) & "@" & drsap(1)
    '            Else
    '                str = "可建單@採購單:0@" & drsap(0) & "@" & drsap(1)
    '            End If
    '            DDLRecord.Items.Add(str)
    '        Loop
    '    Else
    '        DDLRecord.Items.Add("無任何單據篩出 請Check")
    '    End If
    '    drsap.Close()
    '    connsap.Close()
    'End Sub

    'Sub UMITKWSearch(kw As String)
    '    Dim str As String

    '    SqlCmd = "SELECT T0.itemcode,T0.itemname,T0.u_F7 " &
    '            "FROM OITM T0 " &
    '            "where T0.itemcode like '%" & kw & "%' or T0.itemname like '%" & kw & "%' order by T0.itemcode"
    '    drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
    '    DDLRecord.Items.Clear()
    '    If (drsap.HasRows) Then
    '        DDLRecord.Items.Add("已篩選出單據 請選擇")
    '        Do While (drsap.Read())
    '            If (drsap(2) = 0) Then
    '                str = "未建立@" & drsap(0) & "@" & drsap(1)
    '            Else
    '                str = "已建立@" & drsap(0) & "@" & drsap(1)
    '            End If
    '            DDLRecord.Items.Add(str)
    '        Loop
    '    Else
    '        DDLRecord.Items.Add("無任何單據篩出 請Check")
    '    End If
    '    drsap.Close()
    '    connsap.Close()
    '    'BtnDisplay.Enabled = True
    'End Sub


    'Sub CreateUIQTFilterFieldItem()

    '    SqlCmd = "SELECT distinct T0.u_itemcode FROM dbo.[@UMIT] T0 order by T0.u_itemcode"
    '    drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
    '    If (drsap.HasRows) Then
    '        DDLID.Items.Clear()
    '        DDLID.Items.Add("選擇料號")
    '        Do While (drsap.Read())
    '            DDLID.Items.Add(drsap(0))
    '        Loop
    '    End If
    '    drsap.Close()
    '    connsap.Close()

    '    SqlCmd = "SELECT distinct T0.u_vender FROM dbo.[@UIQT] T0 order by T0.u_vender"
    '    drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
    '    If (drsap.HasRows) Then
    '        DDLVender.Items.Clear()
    '        DDLVender.Items.Add("廠商")
    '        Do While (drsap.Read())
    '            DDLVender.Items.Add(drsap(0))
    '        Loop
    '    End If
    '    drsap.Close()
    '    connsap.Close()

    '    TxtBeginDate.Text = ""

    '    TxtEndDate.Text = Format(Now(), "yyyy/MM/dd")

    '    TxtPOFilter.Text = ""

    '    DDLMtype.Items.Clear()
    '    DDLMtype.Items.Add("進料類別")
    '    DDLMtype.Items.Add("銑件")
    '    DDLMtype.Items.Add("車件")
    '    DDLMtype.Items.Add("市購")
    '    DDLMtype.Items.Add("鈑金")
    '    DDLMtype.Items.Add("骨架")

    '    DDLResult.Items.Clear()
    '    DDLResult.Items.Add("選擇結果")
    '    DDLResult.Items.Add("允收")
    '    DDLResult.Items.Add("拒收")
    '    DDLResult.Items.Add("特採")
    '    DDLResult.Items.Add("其他")

    '    DDLAudit.Items.Clear()
    '    DDLAudit.Items.Add("選狀態")
    '    DDLAudit.Items.Add("檢驗中")
    '    DDLAudit.Items.Add("未審單")
    '    DDLAudit.Items.Add("未結案")
    '    DDLAudit.Items.Add("已結案")
    'End Sub

    Protected Sub BtnUpdate_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If (iqctype = 1 Or iqctype = 2) Then
            If (UMITUpdateFieldCheck() = False) Then
                CommUtil.ShowMsg(Me, "欄位資訊不足,請完善欄位資料")
                Exit Sub
            End If
            UMITUpdate()
        End If
        If (iqctype = 3 Or iqctype = 4) Then
            If (UIQTCheck() = False) Then
                Exit Sub
            End If
            UIQTUpdate()
        End If
    End Sub
    Function UMITUpdateFieldCheck()
        Dim i As Integer
        UMITUpdateFieldCheck = True
        For i = 1 To 14
            If (CType(CT.FindControl("txt_iname_" & i), TextBox).Text <> "" Or
                CType(CT.FindControl("txt_ispec_" & i), TextBox).Text <> "" Or
                CType(CT.FindControl("txt_itool_" & i), TextBox).Text <> "") Then
                If (CType(CT.FindControl("txt_iname_" & i), TextBox).Text = "") Then
                    UMITUpdateFieldCheck = False
                    CommUtil.ShowMsg(Me, "第" & i & "列檢驗項目空白")
                End If
                If (CType(CT.FindControl("txt_ispec_" & i), TextBox).Text = "") Then
                    UMITUpdateFieldCheck = False
                    CommUtil.ShowMsg(Me, "第" & i & "列檢驗規格空白")
                End If
                If (CType(CT.FindControl("txt_itool_" & i), TextBox).Text = "") Then
                    UMITUpdateFieldCheck = False
                    CommUtil.ShowMsg(Me, "第" & i & "列檢驗工具空白")
                End If
                If (IsNumeric(CType(CT.FindControl("txt_ispec_" & i), TextBox).Text)) Then
                    If (TolFieldCheck(i) = False) Then
                        UMITUpdateFieldCheck = False
                    End If
                End If
            End If
        Next
    End Function
    Sub UMITUpdate() 'ttttt
        Dim i, k As Integer
        Dim ucode, targetmit As Long
        Dim itemcode As String
        Dim mapno As Long
        Dim InsertStatus, UpdateStatus, insertflag, createflag As Boolean
        Dim seq, setnum As Integer
        setnum = 1
        If (CType(CT.FindControl("txt_iname_1"), TextBox).Text = "") Then
            CommUtil.ShowMsg(Me, "無驗證項目, 需輸入")
            Exit Sub
        End If
        itemcode = Trim(CT.Rows(1).Cells(1).Text)
        'SqlCmd = "SELECT IsNull(T0.U_F6,0) FROM OITM T0 where T0.itemcode='" & itemcode & "'"
        'drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        'If (drsap.HasRows) Then
        '    drsap.Read()
        '    If (drsap(0) = 0) Then
        '        'SqlCmd = "SELECT IsNull(Max(T0.u_F6),0) from OITM T0"
        '        'drsaphead = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsaphead)
        '        'drsaphead.Read()
        '        'mapno = drsaphead(0) + 1
        '        'drsaphead.Close()
        '        'connsaphead.Close()
        '        genmapnoflag = True
        '    Else
        '        'mapno = drsap(0)
        '        genmapnoflag = False
        '    End If
        'End If
        'drsap.Close()
        'connsap.Close()
        mapno = CT.Rows(0).Cells(3).Text
        SqlCmd = "SELECT IsNull(Max(cast(T0.Code as int)),0) from [dbo].[@UMIT] T0"
        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        drsap.Read()
        ucode = drsap(0)
        drsap.Close()
        connsap.Close()
        GetItemDataIntoStructure()
        If (CTItem(0).row = 99) Then
            createflag = True
        Else
            createflag = False
        End If
        seq = 1
        For i = 1 To 14
            insertflag = True
            k = 0
            Do While (CTItem(k).row <> 99)
                If (CTItem(k).row = i) Then
                    targetmit = CTItem(k).icode
                    insertflag = False
                    Exit Do
                End If
                k = k + 1
            Loop
            If (insertflag) Then '新增資料
                If (CType(CT.FindControl("txt_iname_" & i), TextBox).Text <> "") Then
                    ucode = ucode + 1
                    InsertStatus = InsertUMITRecord(i, ucode, seq, mapno)
                    seq = seq + 1
                End If
            Else '原資料更改
                If (CType(CT.FindControl("txt_iname_" & i), TextBox).Text <> "") Then '修改
                    UpdateStatus = UpdateUMITRecord(i, targetmit, seq)
                    seq = seq + 1
                    If (UpdateStatus) Then

                    End If
                Else '刪除項目
                    SqlCmd = "delete from [dbo].[@UMIT] " &
                    "where code = " & targetmit
                    CommUtil.SqlSapExecute("del", SqlCmd, connsap)
                    connsap.Close()
                End If
            End If
        Next
        If (createflag = True) Then 'kkk
            'If (genmapnoflag = True) Then '把mapno寫入oitm , 並把建立iqc資料設定
            '    SqlCmd = "update oitm set " &
            '        "u_F6= " & mapno & " ,u_F7 = " & setnum & " where itemcode ='" & itemcode & "'"
            'Else
            SqlCmd = "update oitm set " &
                    "u_F7 = " & setnum & " where itemcode ='" & itemcode & "'"
            'End If
            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
            connsap.Close()
            CommUtil.ShowMsg(Me, "新增完成")
        Else
            CommUtil.ShowMsg(Me, "修改完成")
        End If
    End Sub
    Function InsertUMITRecord(i As Integer, ucode As Long, seq As Integer, mapno As Long)
        Dim tiname, tispec, tooluse, tol, itemcode As String
        Dim valueflag As Integer
        tiname = Trim(CType(CT.FindControl("txt_iname_" & CStr(i)), TextBox).Text)
        tispec = Trim(CType(CT.FindControl("txt_ispec_" & CStr(i)), TextBox).Text)
        tol = Trim(CType(CT.FindControl("txt_itol_" & CStr(i)), TextBox).Text)
        tooluse = Trim(CType(CT.FindControl("txt_itool_" & CStr(i)), TextBox).Text)
        itemcode = Trim(CT.Rows(1).Cells(1).Text)
        If (IsNumeric(tispec)) Then
            valueflag = 1
        Else
            valueflag = 2
        End If

        SqlCmd = "insert into [dbo].[@UMIT] (code,name,u_mapno,u_tiseq,u_tiname,u_tispec,u_tooluse,u_itemcode,u_tol,u_valueflag) " &
        "values(" & ucode & "," & ucode & "," & mapno & "," & seq & ",'" & tiname & "','" & tispec & "','" & tooluse & "','" & itemcode & "'," &
                "'" & tol & "'," & valueflag & ")"
        InsertUMITRecord = CommUtil.SqlSapExecute("ins", SqlCmd, connsap)
        connsap.Close()
    End Function
    Function UpdateUMITRecord(i As Integer, targetmit As Long, seq As Integer)
        Dim tiname, tispec, tooluse, tol As String
        Dim valueflag As Integer
        tiname = Trim(CType(CT.FindControl("txt_iname_" & CStr(i)), TextBox).Text)
        tispec = Trim(CType(CT.FindControl("txt_ispec_" & CStr(i)), TextBox).Text)
        tol = Trim(CType(CT.FindControl("txt_itol_" & CStr(i)), TextBox).Text)
        tooluse = Trim(CType(CT.FindControl("txt_itool_" & CStr(i)), TextBox).Text)
        If (IsNumeric(tispec)) Then
            valueflag = 1
        Else
            valueflag = 2
        End If
        SqlCmd = "update [dbo].[@UMIT] set " &
        "u_tiname= '" & tiname & "' , u_tispec= '" & tispec & "', " &
        "u_tooluse= '" & tooluse & "' , u_tiseq= " & seq & ",u_tol='" & tol & "',u_valueflag=" & valueflag & " " &
        "where code = " & targetmit
        UpdateUMITRecord = CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
        connsap.Close()
    End Function
    Function UIQTCheck()
        Dim i As Integer
        UIQTCheck = True
        'MsgBox(CType(CT.FindControl("txt_inamount"), TextBox).Text)
        If (CType(CT.FindControl("txt_inamount"), TextBox).Text = "") Then
            CommUtil.ShowMsg(Me, "已進數量需填")
            UIQTCheck = False
        End If
        If (CType(CT.FindControl("txt_famount"), TextBox).Text = "") Then
            CommUtil.ShowMsg(Me, "已檢數量需填")
            UIQTCheck = False
        End If
        If (CType(CT.FindControl("txt_poamount"), TextBox).Text = "") Then
            CommUtil.ShowMsg(Me, "採購數量需填")
            UIQTCheck = False
        End If
        If (CType(CT.FindControl("rbl_mtype"), RadioButtonList).SelectedIndex < 0) Then
            CommUtil.ShowMsg(Me, "進料種類需選其一")
            UIQTCheck = False
        End If

        'Do While (Sheets(sh).Cells(i, 2) <> "" Or Sheets(sh).Cells(i, 6) <> "" Or Sheets(sh).Cells(i, 18) <> "")
        For i = 1 To 14
            If (CType(CT.FindControl("txt_iname_" & i), TextBox).Text <> "" Or
                CType(CT.FindControl("txt_ispec_" & i), TextBox).Text <> "" Or
                CType(CT.FindControl("txt_itool_" & i), TextBox).Text <> "") Then
                If (CType(CT.FindControl("txt_iname_" & i), TextBox).Text = "") Then
                    UIQTCheck = False
                    CommUtil.ShowMsg(Me, "第" & i & "列檢驗項目空白")
                End If
                If (CType(CT.FindControl("txt_ispec_" & i), TextBox).Text = "") Then
                    UIQTCheck = False
                    CommUtil.ShowMsg(Me, "第" & i & "列檢驗規格空白")
                End If
                If (CType(CT.FindControl("txt_itool_" & i), TextBox).Text = "") Then
                    UIQTCheck = False
                    CommUtil.ShowMsg(Me, "第" & i & "列檢驗工具空白")
                End If
                If (IsNumeric(CType(CT.FindControl("txt_ispec_" & i), TextBox).Text)) Then
                    If (TolFieldCheck(i) = False) Then
                        UIQTCheck = False
                    End If
                End If
            End If
        Next
        'Loop
    End Function
    Function TolFieldCheck(row As Integer)
        Dim tol As String
        Dim ok As Boolean
        Dim str() As String
        ok = True
        tol = CType(CT.FindControl("txt_itol_" & row), TextBox).Text

        If (Left(tol, 1) = "±") Then

        ElseIf (Left(tol, 2) = "+-") Then

        ElseIf (Left(tol, 2) = "-+") Then

        ElseIf (Left(tol, 1) = "+") Then
            str = Split(tol, "+")
            If (UBound(str) = 1) Then

            ElseIf (UBound(str) = 2) Then

            Else
                CommUtil.ShowMsg(Me, "誤差欄位+大於2個")
                ok = False
            End If
        ElseIf (Left(tol, 1) = "-") Then
            str = Split(tol, "-")
            If (UBound(str) = 1) Then

            ElseIf (UBound(str) = 2) Then

            Else
                CommUtil.ShowMsg(Me, "誤差欄位-大於2個")
                ok = False
            End If
        Else
            CommUtil.ShowMsg(Me, "誤差欄位沒發現+,-號 ,或第一字元不是+,-或± , 請check")
            ok = False
        End If
        TolFieldCheck = ok
    End Function

    Sub GetItemDataIntoStructure()
        Dim i As Integer
        If (iqctype <> 3 And iqctype <> 5) Then 'item由主檔導入 , 不能把code帶進去,否則會被視為不新增, 是修改而導致錯誤
            If (iqctype = 1 Or iqctype = 2) Then 'MIT(主檔) 顯示 , 建立
                SqlCmd = "SELECT T0.Code " &
                            "FROM dbo.[@UMIT] T0 " &
                            "where T0.u_itemcode='" & itemcode_back & "' order by T0.u_tiseq"
            ElseIf (iqctype = 4) Then 'IQT 顯示 , 建立
                SqlCmd = "SELECT T0.code " &
                            "FROM dbo.[@UIQI] T0 " &
                            "where T0.u_iqtdoc=" & num_back & " order by T0.u_tiseq"
            End If
            drsapitem = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsapitem)
            i = 1
            If (drsapitem.HasRows) Then

                Do While (drsapitem.Read())
                    CTItem(i - 1).row = i
                    CTItem(i - 1).icode = drsapitem(0)
                    CTItem(i).row = 99
                    i = i + 1
                Loop
            End If
            drsapitem.Close()
            connsapitem.Close()
        End If
    End Sub
    Sub UIQTUpdate()
        '以下為本次insert各table 之key field
        'UIQT :ucode
        'UIQI :ucode (docnum)
        'UIQIF:iqiucode
        Dim docnum As Long
        Dim i As Integer
        Dim j, k As Integer
        Dim tiseq As Integer
        Dim rseq As Integer
        Dim targetiqi As Long
        Dim iqiucode As Long
        Dim maxiqifucode As Long
        Dim updiqifucode As Long
        Dim InsertStatus, UpdateStatus, insertflag As Boolean
        If (CType(CT.FindControl("txt_iname_1"), TextBox).Text = "") Then
            CommUtil.ShowMsg(Me, "無驗證項目, 需輸入")
            Exit Sub
        End If
        If (CT.Rows(0).Cells(1).Text <> "") Then
            docnum = CLng(CT.Rows(0).Cells(1).Text)
        Else
            docnum = 0
        End If

        If (docnum = 0) Then '新增
            If (UIQTIniInsertFieldCheck() = False) Then
                CommUtil.ShowMsg(Me, "欄位資訊不足,請完善欄位資料")
                Exit Sub
            End If
            UIQTInsert()
            CommUtil.ShowMsg(Me, "新增完成")
        Else ' update or delete
            '表頭Update
            UpdateUIQTHead(docnum)
            SqlCmd = "SELECT IsNull(Max(cast(T0.Code as int)),0) from [dbo].[@UIQI] T0"
            drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            drsap.Read()
            iqiucode = drsap(0)
            drsap.Close()
            connsap.Close()

            SqlCmd = "SELECT IsNull(Max(cast(T0.Code as int)),0) from [dbo].[@UIQIF] T0"
            drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            drsap.Read()
            maxiqifucode = drsap(0)
            drsap.Close()
            connsap.Close()
            tiseq = 1
            GetItemDataIntoStructure()
            '表身驗證項目Update(UIQI) 處理
            'i = 0
            'Do While (CTItem(i).row <> 99)
            '    MsgBox(CTItem(i).icode)
            '    i = i + 1
            'Loop
            'Exit Sub
            For i = 1 To 14
                k = 0
                insertflag = True
                Do While (CTItem(k).row <> 99)
                    If (CTItem(k).row = i) Then
                        targetiqi = CTItem(k).icode
                        insertflag = False
                        Exit Do
                    End If
                    k = k + 1
                Loop
                If (insertflag) Then '新增資料
                    If (CType(CT.FindControl("txt_iname_" & i), TextBox).Text <> "") Then
                        iqiucode = iqiucode + 1
                        InsertStatus = InsertUIQIRecord(i, iqiucode, tiseq, docnum)
                        tiseq = tiseq + 1
                        If (InsertStatus = True) Then
                            '表身驗證項目之檢驗結果插入(UIQIF)
                            For j = 1 To 9
                                If (CType(CT.FindControl("txt_iresult_" & i & "_" & j), TextBox).Text <> "") Then
                                    rseq = j
                                    maxiqifucode = maxiqifucode + 1
                                    InsertStatus = InsertUIQIFRecord(maxiqifucode, iqiucode, rseq, CType(CT.FindControl("txt_iresult_" & i & "_" & j), TextBox).Text, False)
                                    If (InsertStatus = False) Then

                                    End If
                                End If
                            Next j
                        End If
                    End If
                Else '更新
                    UpdateStatus = UpdateUIQIRecord(i, targetiqi, tiseq)
                    tiseq = tiseq + 1
                    '獲得測試項目之原結果field , 以便如下判斷是否是新增或修改(根據seq)
                    For j = 1 To 9
                        rseq = j
                        SqlCmd = "SELECT T0.u_fieldseq,T0.code from [dbo].[@UIQIF] T0 " &
                                    "where T0.u_iqidoc=" & targetiqi & " and T0.u_fieldseq=" & rseq
                        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)

                        If (CType(CT.FindControl("txt_iresult_" & i & "_" & j), TextBox).Text <> "") Then
                            If (drsap.HasRows) Then '本來有 , 現在也有 , 可能是要更新
                                drsap.Read()
                                updiqifucode = drsap(1)
                                UpdateStatus = UpdateUIQIFRecord(updiqifucode, targetiqi, rseq, CType(CT.FindControl("txt_iresult_" & i & "_" & j), TextBox).Text, True)
                            Else '本來無 , 現在有 , 新增
                                maxiqifucode = maxiqifucode + 1
                                InsertStatus = InsertUIQIFRecord(maxiqifucode, targetiqi, rseq, CType(CT.FindControl("txt_iresult_" & i & "_" & j), TextBox).Text, True)
                            End If
                        Else
                            If (drsap.HasRows) Then '本來有 , 現在無==>刪除
                                'delete result item
                                SqlCmd = "delete from [dbo].[@UIQIF] " &
                                    "where u_iqidoc = " & targetiqi & " and u_fieldseq=" & rseq
                                CommUtil.SqlSapExecute("del", SqlCmd, connsaphead)
                                connsaphead.Close()
                            End If
                        End If
                        drsap.Close()
                        connsap.Close()
                    Next
                End If
            Next
            CommUtil.ShowMsg(Me, "修改完成")
        End If
    End Sub

    Sub UpdateUIQTHead(ucode As Long)
        Dim firstqc, amount, tamount, ngamount, mtype, judge As Integer
        Dim po As Long
        Dim itemcode, vender, createdate, cmemo, amemo, inspector, inspecdate, auditor, auditdate As String
        Dim status As String
        Dim inamount, famount As Integer
        If (CType(CT.FindControl("rbl_firstqc"), RadioButtonList).SelectedIndex >= 0) Then
            firstqc = CType(CT.FindControl("rbl_firstqc"), RadioButtonList).SelectedValue
        Else
            firstqc = 0
        End If
        itemcode = Trim(CT.Rows(1).Cells(1).Text)
        'vender = CType(CT.FindControl("ddl_vender"), DropDownList).SelectedValue 
        vender = CType(CT.FindControl("txt_vender"), TextBox).Text
        createdate = Trim(CT.Rows(1).Cells(7).Text)
        amount = CInt(CType(CT.FindControl("txt_poamount"), TextBox).Text)
        po = CLng(CType(CT.FindControl("txt_ponum"), TextBox).Text)
        tamount = CInt(CType(CT.FindControl("txt_tamount"), TextBox).Text)
        ngamount = CInt(CType(CT.FindControl("txt_ngamount"), TextBox).Text)
        inamount = CInt(CType(CT.FindControl("txt_inamount"), TextBox).Text)
        famount = CInt(CType(CT.FindControl("txt_famount"), TextBox).Text)
        If (CType(CT.FindControl("rbl_mtype"), RadioButtonList).SelectedIndex >= 0) Then
            mtype = CType(CT.FindControl("rbl_mtype"), RadioButtonList).SelectedValue
        Else
            mtype = 0
        End If
        If (CType(CT.FindControl("rbl_judge"), RadioButtonList).SelectedIndex >= 0) Then
            judge = CType(CT.FindControl("rbl_judge"), RadioButtonList).SelectedValue
        Else
            judge = 0
        End If

        cmemo = CType(CT.FindControl("txt_cmemo"), TextBox).Text
        amemo = CType(CT.FindControl("txt_amemo"), TextBox).Text
        inspector = CType(CT.FindControl("btn_inspector"), Button).Text
        If (CType(CT.FindControl("txt_inspecdate"), TextBox).Text <> "") Then
            inspecdate = CType(CT.FindControl("txt_inspecdate"), TextBox).Text
        Else
            inspecdate = "1900/1/1"
        End If
        auditor = CType(CT.FindControl("btn_auditor"), Button).Text
        If (CType(CT.FindControl("txt_auditdate"), TextBox).Text <> "") Then
            auditdate = CType(CT.FindControl("txt_auditdate"), TextBox).Text
        Else
            auditdate = "1900/1/1"
        End If
        status = UIQTStatusGet()
        SqlCmd = "update [dbo].[@UIQT] set " &
        "u_po=" & po & ",u_firstqc=" & firstqc & ",u_amount=" & amount & ",u_tamount=" & tamount & "," &
        "u_ngamount=" & ngamount & ",u_mtype=" & mtype & ",u_judge=" & judge & ",u_itemcode='" & itemcode & "', " &
        "u_vender='" & vender & "', u_cdate='" & createdate & "', u_cmemo='" & cmemo & "', u_amemo='" & amemo & "', u_inspector='" & inspector & "', " &
        "u_inspecdate='" & inspecdate & "', u_auditor='" & auditor & "', u_auditdate='" & auditdate & "',u_status='" & status & "'," &
        "u_inamount=" & inamount & ",u_famount=" & famount & " where u_docnum=" & ucode
        CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
        connsap.Close()
    End Sub

    Function UIQTStatusGet()
        UIQTStatusGet = ""
        If (CType(CT.FindControl("btn_inspector"), Button).Text = "簽名") Then
            UIQTStatusGet = "檢驗中"
        ElseIf (CType(CT.FindControl("btn_auditor"), Button).Text = "簽名" And CType(CT.FindControl("btn_inspector"), Button).Text <> "簽名") Then
            UIQTStatusGet = "未審單"
        ElseIf ((CType(CT.FindControl("btn_auditor"), Button).Text = "簽名" And CType(CT.FindControl("btn_inspector"), Button).Text <> "簽名") Or CType(CT.FindControl("btn_inspector"), Button).Text = "簽名") Then '不會執行到這 , 在上面If檢驗中就會match
            UIQTStatusGet = "未結案"
        ElseIf (CType(CT.FindControl("btn_auditor"), Button).Text <> "簽名") Then
            UIQTStatusGet = "已結案"
        Else
            CommUtil.ShowMsg(Me, "IQC 狀態獲得有問題")
        End If
    End Function
    Function UpdateUIQIRecord(i As Integer, iqiucode As Long, seq As Integer)
        Dim tiname, tispec, tooluse, tol As String
        Dim valueflag As Integer
        tiname = CType(CT.FindControl("txt_iname_" & CStr(i)), TextBox).Text
        tispec = CType(CT.FindControl("txt_ispec_" & CStr(i)), TextBox).Text
        tol = CType(CT.FindControl("txt_itol_" & CStr(i)), TextBox).Text
        tooluse = CType(CT.FindControl("txt_itool_" & CStr(i)), TextBox).Text

        If (IsNumeric(tispec)) Then
            valueflag = 1
        Else
            valueflag = 2
        End If

        SqlCmd = "update [dbo].[@UIQI] set " &
        "u_tiname= '" & tiname & "' , u_tispec= '" & tispec & "', " &
        "u_tooluse= '" & tooluse & "' , u_tiseq= " & seq & "," &
        "u_tol='" & tol & "',u_valueflag=" & valueflag & " " &
        "where code = " & iqiucode
        UpdateUIQIRecord = CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
        connsap.Close()
    End Function
    Function UpdateUIQIFRecord(iqifucode As Long, iqidoc As Long, rseq As Integer, fieldresult As String, preconnection As Boolean)

        SqlCmd = "update [dbo].[@UIQIF] set u_iqidoc=" & iqidoc & ",u_fieldseq=" & rseq & ",u_fieldresult='" & fieldresult & "' " &
        "where code=" & iqifucode
        UpdateUIQIFRecord = CommUtil.SqlSapExecute("upd", SqlCmd, connsaprecord)
        connsaprecord.Close()
    End Function

    Function InsertUIQIRecord(i As Integer, iqiucode As Long, seq As Integer, iqtdoc As Long)
        Dim tiname, tispec, tooluse, tol As String
        Dim valueflag As Integer
        tiname = Trim(CType(CT.FindControl("txt_iname_" & CStr(i)), TextBox).Text)
        tispec = Trim(CType(CT.FindControl("txt_ispec_" & CStr(i)), TextBox).Text)
        tol = Trim(CType(CT.FindControl("txt_itol_" & CStr(i)), TextBox).Text)
        tooluse = Trim(CType(CT.FindControl("txt_itool_" & CStr(i)), TextBox).Text)
        If (IsNumeric(tispec)) Then
            valueflag = 1
        Else
            valueflag = 2
        End If
        SqlCmd = "insert into [dbo].[@UIQI] (code,name,u_iqtdoc,u_tiseq,u_tiname,u_tispec,u_tooluse,u_tol,u_valueflag) " &
        "values(" & iqiucode & "," & iqiucode & "," & iqtdoc & "," & seq & ",'" & tiname & "','" & tispec & "','" & tooluse & "'," &
                "'" & tol & "'," & valueflag & ")"
        InsertUIQIRecord = CommUtil.SqlSapExecute("ins", SqlCmd, connsaprecord)
        connsaprecord.Close()
    End Function

    Function InsertUIQIFRecord(iqifucode As Long, iqidoc As Long, rseq As Integer, fieldresult As String, preconnection As Boolean)

        SqlCmd = "insert into [dbo].[@UIQIF] (code,name,u_iqidoc,u_fieldseq,u_fieldresult) " &
        "values(" & iqifucode & "," & iqifucode & "," & iqidoc & "," & rseq & ",'" & fieldresult & "')"
        InsertUIQIFRecord = CommUtil.SqlSapExecute("ins", SqlCmd, connsaprecord)
        connsaprecord.Close()
    End Function
    Sub UIQTInsert()
        '以下為本次insert各table 之key field
        'UIQT :ucode
        'UIQI :ucode (docnum)
        'UIQIF:iqiucode
        Dim i As Integer
        Dim j As Integer
        Dim tiseq As Integer
        Dim rseq As Integer
        Dim ucode As Long
        Dim iqiucode As Long
        Dim iqifucode As Long
        Dim InsertStatus As Boolean

        If (CType(CT.FindControl("txt_iname_1"), TextBox).Text = "") Then
            CommUtil.ShowMsg(Me, "無驗證項目, 只存檔表頭")
            'Exit Sub
        End If

        SqlCmd = "SELECT IsNull(Max(cast(T0.Code as int)),0) from [dbo].[@UIQT] T0"
        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        drsap.Read()
        ucode = drsap(0) + 1
        drsap.Close()
        connsap.Close()

        InsertUIQTHead(ucode)

        '表身驗證項目插入(UIQI)
        SqlCmd = "SELECT IsNull(Max(cast(T0.Code as int)),0) from [dbo].[@UIQI] T0"
        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        drsap.Read()
        iqiucode = drsap(0)
        drsap.Close()
        connsap.Close()

        SqlCmd = "SELECT IsNull(Max(cast(T0.Code as int)),0) from [dbo].[@UIQIF] T0"
        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        drsap.Read()
        iqifucode = drsap(0)
        drsap.Close()
        connsap.Close()
        tiseq = 1
        i = 1
        For i = 1 To 14
            If (CType(CT.FindControl("txt_iname_" & i), TextBox).Text <> "") Then
                iqiucode = iqiucode + 1
                InsertStatus = InsertUIQIRecord(i, iqiucode, tiseq, ucode)
                If (InsertStatus = True) Then
                    '表身驗證項目之檢驗結果插入(UIQIF)
                    For j = 1 To 9
                        If (CType(CT.FindControl("txt_iresult_" & i & "_" & j), TextBox).Text <> "") Then
                            iqifucode = iqifucode + 1
                            rseq = j
                            InsertStatus = InsertUIQIFRecord(iqifucode, iqiucode, rseq, CType(CT.FindControl("txt_iresult_" & i & "_" & j), TextBox).Text, False)
                            If (InsertStatus = False) Then
                                UIQTDelete(3)
                                Exit Sub
                            End If
                        End If
                    Next j
                Else
                    UIQTDelete(2)
                    Exit Sub
                End If
                tiseq = tiseq + 1
            End If
        Next
    End Sub
    Function UIQTIniInsertFieldCheck()
        UIQTIniInsertFieldCheck = True
        If (CT.Rows(0).Cells(1).Text <> "") Then
            CommUtil.ShowMsg(Me, "單號應空白,由系統生成, 請check")
            UIQTIniInsertFieldCheck = False
            Exit Function
        End If
        If (CT.Rows(1).Cells(1).Text = "") Then
            UIQTIniInsertFieldCheck = False
            CommUtil.ShowMsg(Me, "料號欄位不能空白")
        End If
        If (CType(CT.FindControl("txt_inamount"), TextBox).Text = "") Then
            CommUtil.ShowMsg(Me, "已進數量需填")
            UIQTIniInsertFieldCheck = False
        End If
        If (CType(CT.FindControl("txt_famount"), TextBox).Text = "") Then
            CommUtil.ShowMsg(Me, "已檢數量需填")
            UIQTIniInsertFieldCheck = False
        End If
        If (CType(CT.FindControl("txt_poamount"), TextBox).Text = "") Then
            UIQTIniInsertFieldCheck = False
            CommUtil.ShowMsg(Me, "採購數量欄位不能空白")
        End If
        If (CType(CT.FindControl("txt_ponum"), TextBox).Text = "") Then
            UIQTIniInsertFieldCheck = False
            CommUtil.ShowMsg(Me, "PO欄位不能空白, 無PO請填0")
        End If
        If (CType(CT.FindControl("txt_tamount"), TextBox).Text = "") Then
            UIQTIniInsertFieldCheck = False
            CommUtil.ShowMsg(Me, "抽驗數欄位不能空白")
        End If

        If (CType(CT.FindControl("txt_ngamount"), TextBox).Text = "") Then
            UIQTIniInsertFieldCheck = False
            CommUtil.ShowMsg(Me, "不良數欄位不能空白")
        End If

        If (CType(CT.FindControl("rbl_mtype"), RadioButtonList).SelectedIndex < 0) Then
            UIQTIniInsertFieldCheck = False
            CommUtil.ShowMsg(Me, "進料類別一定要選一個")
        End If

        '以下為檢驗項目check
        For i = 1 To 14
            If (CType(CT.FindControl("txt_iname_" & i), TextBox).Text <> "" Or
                CType(CT.FindControl("txt_ispec_" & i), TextBox).Text <> "" Or
                CType(CT.FindControl("txt_itool_" & i), TextBox).Text <> "") Then
                If (CType(CT.FindControl("txt_iname_" & i), TextBox).Text = "") Then
                    UIQTIniInsertFieldCheck = False
                    CommUtil.ShowMsg(Me, "第" & i & "列檢驗項目空白")
                End If
                If (CType(CT.FindControl("txt_ispec_" & i), TextBox).Text = "") Then
                    UIQTIniInsertFieldCheck = False
                    CommUtil.ShowMsg(Me, "第" & i & "列檢驗規格空白")
                End If
                If (CType(CT.FindControl("txt_itool_" & i), TextBox).Text = "") Then
                    UIQTIniInsertFieldCheck = False
                    CommUtil.ShowMsg(Me, "第" & i & "列檢驗工具空白")
                End If
                If (IsNumeric(CType(CT.FindControl("txt_ispec_" & i), TextBox).Text)) Then
                    If (TolFieldCheck(i) = False) Then
                        UIQTIniInsertFieldCheck = False
                    End If
                End If
            End If
        Next
    End Function
    Sub InsertUIQTHead(ucode As Long)
        Dim firstqc, amount, tamount, ngamount, mtype, judge, inamount, famount As Integer
        Dim po As Long
        Dim mapno As Long
        Dim itemcode, vender, createdate, cmemo, amemo, inspector, inspecdate, auditor, auditdate As String
        Dim status As String
        If (CType(CT.FindControl("rbl_firstqc"), RadioButtonList).SelectedIndex >= 0) Then
            firstqc = CType(CT.FindControl("rbl_firstqc"), RadioButtonList).SelectedValue
        Else
            firstqc = 0
        End If
        itemcode = Trim(CT.Rows(1).Cells(1).Text)
        vender = CType(CT.FindControl("txt_vender"), TextBox).Text
        createdate = Trim(CT.Rows(1).Cells(7).Text)
        amount = CInt(CType(CT.FindControl("txt_poamount"), TextBox).Text)
        po = CLng(CType(CT.FindControl("txt_ponum"), TextBox).Text)
        tamount = CInt(CType(CT.FindControl("txt_tamount"), TextBox).Text)
        ngamount = CInt(CType(CT.FindControl("txt_ngamount"), TextBox).Text)
        inamount = CInt(CType(CT.FindControl("txt_inamount"), TextBox).Text)
        famount = CInt(CType(CT.FindControl("txt_famount"), TextBox).Text)
        If (CType(CT.FindControl("rbl_mtype"), RadioButtonList).SelectedIndex >= 0) Then
            mtype = CType(CT.FindControl("rbl_mtype"), RadioButtonList).SelectedValue
        Else
            mtype = 0
        End If
        If (CType(CT.FindControl("rbl_judge"), RadioButtonList).SelectedIndex >= 0) Then
            judge = CType(CT.FindControl("rbl_judge"), RadioButtonList).SelectedValue
        Else
            judge = 0
        End If

        cmemo = CType(CT.FindControl("txt_cmemo"), TextBox).Text
        amemo = CType(CT.FindControl("txt_amemo"), TextBox).Text
        inspector = CType(CT.FindControl("btn_inspector"), Button).Text
        mapno = CLng(Trim(CT.Rows(0).Cells(3).Text))
        inspecdate = CType(CT.FindControl("txt_inspecdate"), TextBox).Text
        auditdate = CType(CT.FindControl("txt_auditdate"), TextBox).Text
        auditor = CType(CT.FindControl("btn_auditor"), Button).Text
        '以下等介面完成再處裡
        'If (Sheets(sh).Cells(27, 15) <> "") Then
        '    inspecdate = Sheets(sh).Cells(27, 15)
        '    inspecdate = Sheets(sh).Cells(27, 15)
        'Else
        '    inspecdate = "1900/1/1"
        'End If
        'auditor = Sheets(sh).Cells(25, 17)
        'If (Sheets(sh).Cells(27, 17) <> "") Then
        '    auditdate = Sheets(sh).Cells(27, 17)
        'Else
        '    auditdate = "1900/1/1"
        'End If
        status = UIQTStatusGet()
        SqlCmd = "insert into [dbo].[@UIQT] (code,name,u_docnum,u_po,u_firstqc,u_amount,u_tamount,u_ngamount,u_mtype,u_judge,u_itemcode, " &
        "u_vender, u_cdate, u_cmemo, u_amemo, u_inspector, u_inspecdate, u_auditor, u_auditdate,u_mapno,u_status,u_inamount,u_famount) " &
        "values(" & ucode & "," & ucode & "," & ucode & "," & po & "," & firstqc & "," & amount & "," & tamount & "," & ngamount & "," & mtype & "," & judge & ",'" & itemcode & "', " &
        "'" & vender & "','" & createdate & "','" & cmemo & "','" & amemo & "','" & inspector & "','" & inspecdate & "','" & auditor & "','" & auditdate & "'," & mapno & ",'" & status & "'" &
        "," & inamount & "," & famount & ")"
        CommUtil.SqlSapExecute("ins", SqlCmd, connsap)
        connsap.Close()

    End Sub

    Protected Sub BtnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If (iqctype = 1 Or iqctype = 2) Then
            UMITDelete()
        ElseIf (iqctype = 3 Or iqctype = 4) Then
            UIQTDelete(3)
        End If
        Response.Redirect("qc.aspx?smid=qc&smode=0&funindex=" & Request.QueryString("funindex") & "&indexpage=" & Request.QueryString("indexpage"))
    End Sub

    Sub UMITDelete()
        Dim itemcode As String
        itemcode = CT.Rows(1).Cells(1).Text
        SqlCmd = "delete from [dbo].[@UMIT] " &
            "where u_itemcode = '" & itemcode & "'"
        CommUtil.SqlSapExecute("del", SqlCmd, connsap)
        connsap.Close()

        SqlCmd = "update oitm set " &
            "u_F7 = 0 where itemcode ='" & itemcode & "'"
        CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
        connsap.Close()
        CommUtil.ShowMsg(Me, "刪除完成")
    End Sub

    Sub UIQTDelete(dtype As Integer) ' 1: 刪IQT  2:刪IQT及IQI  3:刪IQT IQI IQIF
        Dim docnum As Long
        Dim delflag As Boolean
        docnum = CLng(CT.Rows(0).Cells(1).Text)
        If (dtype = 3) Then
            SqlCmd = "SELECT T0.code " &
            "FROM dbo.[@UIQI] T0 " &
            "where T0.u_iqtdoc=" & docnum
            drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (drsap.HasRows) Then
                Do While (drsap.Read())
                    SqlCmd = "delete from [dbo].[@UIQIF] " &
                    "where u_iqidoc = " & drsap(0)
                    delflag = CommUtil.SqlSapExecute("del", SqlCmd, connsaphead)
                    connsaphead.Close()
                Loop
            End If
            drsap.Close()
            connsap.Close()
        End If
        If (dtype = 2 Or dtype = 3) Then
            SqlCmd = "delete from [dbo].[@UIQI] " &
                "where u_iqtdoc = " & docnum
            CommUtil.SqlSapExecute("del", SqlCmd, connsaphead)
            connsaphead.Close()
        End If

        If (dtype = 1 Or dtype = 2 Or dtype = 3) Then
            SqlCmd = "delete from [dbo].[@UIQT] " &
                "where code = " & docnum
            CommUtil.SqlSapExecute("del", SqlCmd, connsaphead)
            connsaphead.Close()
        End If
        CommUtil.ShowMsg(Me, "刪除完畢")
    End Sub
    Protected Sub BtnInspector_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        If (sender.Text = "簽名") Then
            SqlCmd = "SELECT u_name from OUSR T0 where user_code='" & Session("sapid") & "'"
            drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (drsap.HasRows) Then
                drsap.Read()
                sender.Text = drsap(0)
                CType(CT.FindControl("txt_inspecdate"), TextBox).Text = Format(Now(), "yyyy/MM/dd")
            Else
                CommUtil.ShowMsg(Me, "SAP無此" & Session("sapid") & "帳號存在, 請Check")
            End If
        Else
            sender.Text = "簽名"
            CType(CT.FindControl("txt_inspecdate"), TextBox).Text = ""
        End If
    End Sub

    Protected Sub BtnAuditor_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If (sender.Text = "簽名") Then
            SqlCmd = "SELECT u_name from OUSR T0 where user_code='" & Session("sapid") & "'"
            drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (drsap.HasRows) Then
                drsap.Read()
                sender.Text = drsap(0)
                CType(CT.FindControl("txt_auditdate"), TextBox).Text = Format(Now(), "yyyy/MM/dd")
            Else
                CommUtil.ShowMsg(Me, "SAP無此" & Session("sapid") & "帳號存在, 請Check")
            End If
        Else
            sender.Text = "簽名"
            CType(CT.FindControl("txt_auditdate"), TextBox).Text = ""
        End If
    End Sub

    Protected Sub BtnDraw_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim OpenFileDialog1 As New Windows.Forms.OpenFileDialog With {
            .CheckFileExists = True,
            .Filter = "pdf files (*.pdf)|*.*",
            .InitialDirectory = "C:\QC\圖檔\",
            .Multiselect = False
        }
        'Dim invokeThread As Threading.Thread
        'invokeThread = New Threading.Thread(New Threading.ThreadStart(AddressOf InvokeMethod))
        'invokeThread.SetApartmentState(Threading.ApartmentState.STA)
        'invokeThread.Start()
        'invokeThread.Join()
        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then '需在aspx檔中@ Page後加 ASPCompat="true" 才能show出dialog 或如上加入額外執行緒
            '把OpenFileDialog1.ShowDialog() 放入 Sub InvokeMethod 中 , 如下
            'Function InvokeMethod()
            '    InvokeMethod = OpenFileDialog1.ShowDialog()
            'End Function

            Process.Start(OpenFileDialog1.FileName)
        End If
    End Sub

    'Protected Sub BtnDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    '    'MsgBox(iqctype & "--" & cflag & "---" & mode)

    'End Sub

    Protected Sub BtnUpload_Click(ByVal sender As Object, ByVal e As System.EventArgs) 'kkkkk
        Dim targetPath, targetLocal, act As String
        Dim filename, orgname As String
        Dim di As DirectoryInfo
        act = ""
        filename = CT.Rows(0).Cells(3).Text & "." & FileUL.FileName
        targetLocal = Application("localdir") & "QC\DW\" & filename
        targetPath = HttpContext.Current.Server.MapPath("~/") & "\AttachFile\QC\DW\" & filename
        di = New DirectoryInfo(HttpContext.Current.Server.MapPath("~/") & "\AttachFile\QC\DW\")
        Dim fi As FileInfo() = di.GetFiles(CT.Rows(0).Cells(3).Text & ".*")
        If (fi.Length = 0) Then
            orgname = ""
        Else
            orgname = fi(0).Name
        End If
        If (sender.Text = "上傳圖檔") Then
            If (FileUL.HasFile) Then
                If (File.Exists(HttpContext.Current.Server.MapPath("~/") & "\AttachFile\QC\DW\" & orgname)) Then
                    IO.File.Delete(Application("localdir") & "QC\DW\" & orgname)
                    IO.File.Delete(HttpContext.Current.Server.MapPath("~/") & "\AttachFile\QC\DW\" & orgname)
                End If
                FileUL.SaveAs(targetPath)
                FileUL.SaveAs(targetLocal)
                act = "fileupload"
            Else
                CommUtil.ShowMsg(Me, "未指定上傳檔案")
            End If
        ElseIf (sender.Text = "刪除圖檔") Then
            IO.File.Delete(Application("localdir") & "QC\DW\" & orgname)
            IO.File.Delete(HttpContext.Current.Server.MapPath("~/") & "\AttachFile\QC\DW\" & orgname)
            act = "filedelete"
        End If
        If (act = "fileupload" Or act = "filedelete") Then
            If (Request.QueryString("funindex") = 1 Or Request.QueryString("funindex") = 2) Then
                Response.Redirect("iqc.aspx?smid=qc&smode=1&act=" & act & "&iqctype=" & iqctype & "&funindex=" & Request.QueryString("funindex") &
            "&indexpage=" & Request.QueryString("indexpage") &
            "&itemcode=" & itemcode_back &
            "&itemname=" & itemname_back & "&mode=" & mode)
            ElseIf (Request.QueryString("funindex") = 3) Then
                Response.Redirect("iqc.aspx?smid=qc&smode=1&act=" & act & "&iqctype=" & iqctype & "&funindex=" & Request.QueryString("funindex") &
            "&indexpage=" & Request.QueryString("indexpage") &
            "&itemcode=" & itemcode_back &
            "&itemname=" & itemname_back & "&po=" & ponum_back & "&mode=" & mode)
            ElseIf (Request.QueryString("funindex") = 4) Then
                Response.Redirect("iqc.aspx?smid=qc&smode=1&act=" & act & "&iqctype=" & iqctype & "&funindex=" & Request.QueryString("funindex") &
            "&indexpage=" & Request.QueryString("indexpage") &
            "&itemname=" & itemname_back & "&docnum=" & num_back & "&mode=" & mode)
            End If
        End If
    End Sub

    Sub DisplayIQCData()
        If (iqctype <> 0) Then
            If (iqctype = 1) Then 'MIT(主檔) 顯示
                itemcode_back = ViewState("itemcode_back")
                itemname_back = ViewState("itemname_back")
                SqlCmd = "SELECT T0.itemname,T0.u_F6 " &
                        "FROM OITM T0 " &
                        "where T0.itemcode='" & itemcode_back & "'"
            ElseIf (iqctype = 2) Then 'MIT(主檔) 建立
                itemcode_back = ViewState("itemcode_back")
                itemname_back = ViewState("itemname_back")
                SqlCmd = "SELECT T0.itemname,IsNull(T0.u_F6,0) " &
                        "FROM OITM T0 " &
                        "where T0.itemcode='" & itemcode_back & "'"
            ElseIf (iqctype = 3 Or iqctype = 5) Then '新建IQT
                ponum_back = ViewState("ponum_back")
                itemcode_back = ViewState("itemcode_back")
                itemname_back = ViewState("itemname_back")
                'If (iqctype = 5) Then
                rest_inamount = ViewState("rest_inamount")
                po_amount = ViewState("po_amount")
                'End If
                SqlCmd = "SELECT T1.quantity,T0.cardname,T2.u_F6 " &
                        "FROM OPOR T0 INNER JOIN POR1 T1 ON T0.DocEntry = T1.DocEntry INNER JOIN OITM T2 ON T1.ItemCode = T2.ItemCode " &
                        "where T0.docnum=" & ponum_back & " and T1.itemcode='" & itemcode_back & "'"
            ElseIf (iqctype = 4) Then '顯示IQT
                num_back = ViewState("num_back")
                itemname_back = ViewState("itemname_back")
                SqlCmd = "SELECT T0.u_inamount,T0.u_famount,T0.u_docnum,T0.u_itemcode,T0.u_firstqc,T0.u_vender, " &
                        "T0.u_cdate,T0.u_amount,T0.u_po,T0.u_tamount,T0.u_ngamount,T0.u_mtype,T0.u_judge, " &
                        "T0.u_cmemo,T0.u_amemo,T0.u_inspector,IsNull(T0.u_inspecdate,''), " &
                        "T0.u_auditor,IsNull(T0.u_auditdate,''),IsNull(T0.u_mapno,0) " &
                        "FROM dbo.[@UIQT] T0 " &
                        "where T0.u_docnum=" & num_back
            End If
            drsaphead = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsaphead)
            If (drsaphead.HasRows) Then
                drsaphead.Read()
                PutIQCData(mode)
            Else
                If (iqctype = 1 Or iqctype = 2) Then
                    CommUtil.ShowMsg(Me, "在SAP查無" & itemcode_back)
                ElseIf (iqctype = 3) Then
                    CommUtil.ShowMsg(Me, "在SAP查無" & ponum_back & "之採購單")
                ElseIf (iqctype = 4) Then
                    CommUtil.ShowMsg(Me, "查無" & num_back & "之IQC單")
                End If
            End If
            drsaphead.Close()
            connsaphead.Close()
        Else
            PutIQCData(mode)
        End If
    End Sub
    Sub PutIQCData(showmode As String)
        If (showmode = "showempty") Then
            showmode = "showvalue"
        End If
        PutCTHead1(showmode)
        PutCTHead2(showmode)
        PutCTHead3(showmode)
        PutCTRecord(showmode)
        PutCTTail_2(showmode)
        PutCTTail_3(showmode)
        PutCTTail_4(showmode)
        If (iqctype = 3 Or iqctype = 4) Then
            If (RBLMode.SelectedIndex = 0) Then
                CType(CT.FindControl("rbl_firstqc"), RadioButtonList).Enabled = False
                CType(CT.FindControl("rbl_mtype"), RadioButtonList).Enabled = False
                CType(CT.FindControl("rbl_judge"), RadioButtonList).Enabled = False
            Else
                CType(CT.FindControl("rbl_firstqc"), RadioButtonList).Enabled = True
                CType(CT.FindControl("rbl_mtype"), RadioButtonList).Enabled = True
                CType(CT.FindControl("rbl_judge"), RadioButtonList).Enabled = True
            End If
        End If
    End Sub

    Sub ResultJudge(i As Integer, j As Integer, tol As String, stdval As String, actualvalue As String, showmode As String)
        Dim hivalue, lowvalue As Double
        Dim str() As String
        If (IsNumeric(stdval) And actualvalue <> "") Then
            stdval = CDbl(stdval)
            If (IsNumeric(actualvalue)) Then
                actualvalue = CDbl(actualvalue)
            Else
                CommUtil.ShowMsg(Me, actualvalue & "不是數字")
            End If
            If (Left(tol, 1) = "±") Then
                hivalue = stdval + CDbl(Mid(tol, 2, 5))
                lowvalue = stdval - CDbl(Mid(tol, 2, 5))
            ElseIf (Left(tol, 2) = "+-") Then
                hivalue = stdval + CDbl(Mid(tol, 3, 5))
                lowvalue = stdval - CDbl(Mid(tol, 3, 5))
            ElseIf (Left(tol, 2) = "-+") Then
                hivalue = stdval + CDbl(Mid(tol, 3, 5))
                lowvalue = stdval - CDbl(Mid(tol, 3, 5))
            ElseIf (Left(tol, 1) = "+") Then
                str = Split(tol, "+")
                If (UBound(str) = 1) Then
                    hivalue = stdval + CDbl(Mid(tol, 2, 5))
                    lowvalue = stdval
                Else
                    hivalue = stdval + CDbl(str(2))
                    lowvalue = stdval + CDbl(str(1))
                End If
            ElseIf (Left(tol, 1) = "-") Then
                str = Split(tol, "-")
                If (UBound(str) = 1) Then
                    hivalue = stdval
                    lowvalue = stdval - CDbl(Mid(tol, 2, 5))
                Else
                    hivalue = stdval - CDbl(str(1))
                    lowvalue = stdval - CDbl(str(2))
                End If
            End If
            lowvalue = lowvalue - 0.000001
            If (actualvalue >= lowvalue And actualvalue <= hivalue) Then
                If (showmode = "edit") Then
                    CType(CT.FindControl("txt_iresult_" & i & "_" & j), TextBox).BackColor = Drawing.Color.Chartreuse
                    CT.Rows(i + 4).Cells(j + 3).BackColor = Drawing.Color.White
                Else
                    CT.Rows(i + 4).Cells(j + 3).BackColor = Drawing.Color.Chartreuse
                End If
            Else
                If (showmode = "edit") Then
                    CType(CT.FindControl("txt_iresult_" & i & "_" & j), TextBox).BackColor = Drawing.Color.Red
                    If (actualvalue < lowvalue) Then
                        CType(CT.FindControl("txt_iresult_" & i & "_" & j), TextBox).ToolTip = "比下限值低:" & CStr(CInt(1000 * (lowvalue - actualvalue))) & "um"
                    Else
                        CType(CT.FindControl("txt_iresult_" & i & "_" & j), TextBox).ToolTip = "比上限值高:" & CStr(CInt(1000 * (actualvalue - hivalue))) & "um"
                    End If
                    CT.Rows(i + 4).Cells(j + 3).BackColor = Drawing.Color.White
                Else
                    CT.Rows(i + 4).Cells(j + 3).BackColor = Drawing.Color.Red
                    If (actualvalue < lowvalue) Then
                        CT.Rows(i + 4).Cells(j + 3).ToolTip = "比下限值低:" & CStr(CInt(1000 * (lowvalue - actualvalue))) & "um"
                    Else
                        CT.Rows(i + 4).Cells(j + 3).ToolTip = "比上限值高:" & CStr(CInt(1000 * (actualvalue - hivalue))) & "um"
                    End If
                End If
            End If
        Else
            If (showmode = "edit") Then
                If (iqctype = 3 Or iqctype = 4) Then
                    'If (preddlfun <> 1) Then
                    CType(CT.FindControl("txt_iresult_" & i & "_" & j), TextBox).BackColor = Drawing.Color.White
                    'End If
                End If
                CT.Rows(i + 4).Cells(j + 3).BackColor = Drawing.Color.White
            Else
                CT.Rows(i + 4).Cells(j + 3).BackColor = Drawing.Color.White
            End If
        End If
    End Sub
    Function GetItemHiLoLimit(tol As String, stdval As String)
        Dim hivalue, lowvalue As Double
        Dim str() As String
        stdval = CDbl(stdval)
        If (Left(tol, 1) = "±") Then
            hivalue = stdval + CDbl(Mid(tol, 2, 5))
            lowvalue = stdval - CDbl(Mid(tol, 2, 5))
        ElseIf (Left(tol, 2) = "+-") Then
            hivalue = stdval + CDbl(Mid(tol, 3, 5))
            lowvalue = stdval - CDbl(Mid(tol, 3, 5))
        ElseIf (Left(tol, 2) = "-+") Then
            hivalue = stdval + CDbl(Mid(tol, 3, 5))
            lowvalue = stdval - CDbl(Mid(tol, 3, 5))
        ElseIf (Left(tol, 1) = "+") Then
            str = Split(tol, "+")
            If (UBound(str) = 1) Then
                hivalue = stdval + CDbl(Mid(tol, 2, 5))
                lowvalue = stdval
            Else
                hivalue = stdval + CDbl(str(2))
                lowvalue = stdval + CDbl(str(1))
            End If
        ElseIf (Left(tol, 1) = "-") Then
            str = Split(tol, "-")
            If (UBound(str) = 1) Then
                hivalue = stdval
                lowvalue = stdval - CDbl(Mid(tol, 2, 5))
            Else
                hivalue = stdval - CDbl(str(1))
                lowvalue = stdval - CDbl(str(2))
            End If
        End If
        GetItemHiLoLimit = CStr(lowvalue) & "~" & CStr(hivalue)
    End Function
    Sub PutCTRecord(showmode As String)
        Dim iname, ispec, itol, iresult, itool As String
        Dim i, j As Integer
        Dim lastitem, lastrecord, nextrecord As Boolean
        Dim iid_num, rid_num As String
        lastitem = False
        'If (mode <> "showempty") Then '在page load 那已確認不會是showempty
        If (iqctype <> 0) Then
            If (iqctype = 1 Or iqctype = 2 Or iqctype = 3 Or iqctype = 5) Then 'MIT(主檔) 顯示(1) , 建立(2),IQC新建(3) ,IQC二次建(5)
                SqlCmd = "SELECT T0.u_tiname,T0.u_tispec,T0.u_tol,T0.u_tooluse,T0.Code " &
                        "FROM dbo.[@UMIT] T0 " &
                        "where T0.u_itemcode='" & itemcode_back & "' order by T0.u_tiseq"
            ElseIf (iqctype = 4) Then 'IQT 顯示 
                SqlCmd = "SELECT T0.u_tiname,T0.u_tispec,T0.u_tol,T0.u_tooluse,T0.code " &
                        "FROM dbo.[@UIQI] T0 " &
                        "where T0.u_iqtdoc=" & num_back & " order by T0.u_tiseq"
            End If
            drsapitem = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsapitem)
        Else

        End If
        'End If

        'If (drsapitem.HasRows) Then
        For i = 1 To 14
            If (iqctype <> 0) Then
                If (drsapitem.HasRows) Then
                    If (drsapitem.Read() And lastitem = False) Then
                        iname = drsapitem(0)
                        ispec = drsapitem(1)
                        itol = drsapitem(2)
                        itool = drsapitem(3)
                        iid_num = CStr(i)
                        If (iqctype <> 3 And iqctype <> 5) Then '此為主檔導入 ,不能視為已存在
                            CTItem(i - 1).row = i
                            CTItem(i - 1).icode = drsapitem(4)
                            CTItem(i).row = 99
                        End If
                    Else
                        lastitem = True
                        iname = ""
                        ispec = ""
                        itol = ""
                        itool = ""
                        iid_num = CStr(i)
                    End If
                Else
                    lastitem = True
                    iname = ""
                    ispec = ""
                    itol = ""
                    itool = ""
                    iid_num = CStr(i)
                End If
            Else
                iname = ""
                ispec = ""
                itol = ""
                itool = ""
                iid_num = CStr(i)
            End If
            CT.Rows(i + 4).Cells(0).Text = i '內容從第5列開始
            If (showmode = "edit") Then
                CType(CT.FindControl("txt_iname_" & iid_num), TextBox).Text = iname
                CType(CT.FindControl("txt_ispec_" & iid_num), TextBox).Text = ispec
                CType(CT.FindControl("txt_itol_" & iid_num), TextBox).Text = itol
                If (IsNumeric(ispec)) Then
                    CType(CT.FindControl("txt_ispec_" & iid_num), TextBox).ToolTip = GetItemHiLoLimit(itol, ispec)
                End If
            ElseIf (showmode = "showvalue") Then
                CT.Rows(i + 4).Cells(1).Text = iname
                CT.Rows(i + 4).Cells(2).Text = ispec
                CT.Rows(i + 4).Cells(3).Text = itol
                If (IsNumeric(ispec)) Then
                    CT.Rows(i + 4).Cells(2).ToolTip = GetItemHiLoLimit(itol, ispec)
                End If
            End If
            If (iqctype = 4 And lastitem = False) Then 'IQT 顯示
                SqlCmd = "SELECT T0.u_fieldresult ,T0.u_fieldseq,T0.code " &
                             "FROM dbo.[@UIQIF] T0 " &
                             "where T0.u_iqidoc=" & drsapitem(4) & " order by T0.u_fieldseq"
                drsaprecord = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsaprecord)
                If (drsaprecord.HasRows) Then
                    lastrecord = False
                    nextrecord = True
                Else
                    lastrecord = True
                    nextrecord = False
                End If
            Else
                lastrecord = False
                nextrecord = True
            End If

            For j = 1 To 9
                rid_num = CStr(i) & "_" & CStr(j)
                If (iqctype = 4 And lastitem = False) Then 'IQT 顯示
                    If (nextrecord) Then
                        If (Not drsaprecord.Read()) Then
                            lastrecord = True
                        End If
                    End If
                    If (lastrecord = False) Then
                        If (drsaprecord(1) = j) Then
                            iresult = drsaprecord(0)
                            'rid_num = rid_num & "_" & drsaprecord(2)
                            nextrecord = True
                            'If (IsNumeric(ispec)) Then
                            'ResultJudge(i, j, itol, ispec, iresult, showmode)
                            'End If
                        Else
                            iresult = ""
                            nextrecord = False
                        End If
                    Else
                        iresult = ""
                    End If
                Else
                    iresult = ""
                End If

                If (showmode = "edit") Then
                    CType(CT.FindControl("txt_iresult_" & rid_num), TextBox).Text = iresult
                ElseIf (showmode = "showvalue") Then
                    CT.Rows(i + 4).Cells(j + 3).Text = iresult
                End If
                ResultJudge(i, j, itol, ispec, iresult, showmode)
            Next
            If (showmode = "edit") Then
                CType(CT.FindControl("txt_itool_" & iid_num), TextBox).Text = itool
            ElseIf (showmode = "showvalue") Then
                CT.Rows(i + 4).Cells(j + 3).Text = itool '此時j應=10
            End If
            If (iqctype = 4 And lastitem = False) Then
                drsaprecord.Close()
                connsaprecord.Close()
            End If
        Next
        'End If
        If (iqctype <> 0) Then
            drsapitem.Close()
            connsapitem.Close()
        End If
    End Sub

    Sub PutCTHead1(showmode As String)
        Dim mapno, docnum As String
        Dim inamount, famount As String
        Dim firstqc As Integer
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        firstqc = 0
        inamount = ""
        famount = ""
        mapno = ""
        docnum = ""
        If (showmode <> "showempty") Then
            If (iqctype = 1) Then 'MIT 顯示
                firstqc = 0
                inamount = ""
                famount = ""
                mapno = CStr(drsaphead(1))
                docnum = ""
            ElseIf (iqctype = 2) Then 'MIT 建立
                firstqc = 0
                inamount = ""
                famount = ""
                If (drsaphead(1) = 0) Then
                    SqlCmd = "SELECT IsNull(Max(T0.u_F6),0) from OITM T0"
                    drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                    drL.Read()
                    mapno = drL(0) + 1
                    drL.Close()
                    connL.Close()
                    SqlCmd = "update oitm set " &
                    "u_F6 = " & mapno & " where itemcode ='" & itemcode_back & "'"
                    CommUtil.SqlSapExecute("upd", SqlCmd, connL)
                    connL.Close()
                Else
                    mapno = CStr(drsaphead(1))
                End If
                docnum = ""
            ElseIf (iqctype = 3) Then 'IQT新建立
                firstqc = 0
                inamount = rest_inamount
                famount = ""
                mapno = CStr(drsaphead(2))
                docnum = ""
            ElseIf (iqctype = 4) Then 'IQT 顯示
                firstqc = CStr(drsaphead(4))
                inamount = CStr(drsaphead(0))
                famount = CStr(drsaphead(1))
                mapno = CStr(drsaphead(19))
                docnum = CStr(drsaphead(2))
            ElseIf (iqctype = 5) Then 'IQT續建立
                firstqc = 0
                inamount = rest_inamount
                famount = ""
                mapno = CStr(drsaphead(2))
                docnum = ""
            End If
        End If
        CT.Rows(0).Cells(1).Text = docnum
        CT.Rows(0).Cells(3).Text = mapno
        'If (mapno = "") Then
        'BtnDraw.Enabled = False
        'Else
        'BtnDraw.Enabled = True
        'End If
        'BtnDraw.Text = "開啟" & CStr(mapno) & "*.pdf檔"
        If (showmode = "edit") Then
            CType(CT.FindControl("txt_inamount"), TextBox).Text = inamount
            CType(CT.FindControl("txt_famount"), TextBox).Text = famount
        ElseIf (showmode = "showvalue") Then
            CT.Rows(0).Cells(5).Text = inamount
            CT.Rows(0).Cells(7).Text = famount
        End If
        If (iqctype = 3 Or iqctype = 4) Then
            If (firstqc <> 0) Then
                CType(CT.FindControl("rbl_firstqc"), RadioButtonList).SelectedValue = firstqc
            Else
                CType(CT.FindControl("rbl_firstqc"), RadioButtonList).SelectedIndex = 0
            End If
        Else

        End If
    End Sub

    Sub PutCTHead2(showmode As String)
        Dim itemcode, itemname, vender, createdate As String
        itemcode = ""
        itemname = ""
        vender = ""
        createdate = ""
        If (showmode <> "showempty") Then
            If (iqctype = 1 Or iqctype = 2) Then 'MIT 顯示 , 建立
                itemcode = itemcode_back
                itemname = itemname_back
                vender = ""
                createdate = ""
            ElseIf (iqctype = 3 Or iqctype = 5) Then 'IQT建立
                itemcode = itemcode_back
                itemname = itemname_back
                vender = drsaphead(1)
                createdate = Format(Now(), "yyyy/MM/dd")
            ElseIf (iqctype = 4) Then 'IQT 顯示
                itemcode = drsaphead(3)
                itemname = itemname_back
                vender = drsaphead(5)
                createdate = drsaphead(6)
            End If
        End If
        CT.Rows(1).Cells(1).Text = itemcode
        CT.Rows(1).Cells(3).Text = itemname
        CT.Rows(1).Cells(7).Text = createdate
        If (showmode = "edit") Then
            CType(CT.FindControl("txt_vender"), TextBox).Text = vender
        ElseIf (showmode = "showvalue") Then
            CT.Rows(1).Cells(5).Text = vender
        End If
    End Sub

    Sub PutCTHead3(showmode As String)
        Dim po As String
        Dim poamount, tamount, ngamount As String
        Dim mtype As Integer
        tamount = ""
        ngamount = ""
        mtype = 0
        po = ""
        poamount = ""
        If (showmode <> "showempty") Then
            If (iqctype = 1 Or iqctype = 2) Then 'MIT 顯示 , 建立
                'nothing
            ElseIf (iqctype = 3) Then 'IQT建立
                tamount = CInt(rest_inamount) 'CStr(CInt(drsaphead(0)))
                po = ponum_back
                poamount = po_amount 'CStr(CInt(drsaphead(0)))
            ElseIf (iqctype = 4) Then 'IQT 顯示
                tamount = CStr(CInt(drsaphead(9)))
                ngamount = CStr(drsaphead(10))
                mtype = drsaphead(11)
                po = CStr(drsaphead(8))
                poamount = CStr(CInt(drsaphead(7)))
                'ElseIf (iqctype = 5) Then 'IQT建立
                '   tamount = rest_inamount
                '  po = ponum_back
                ' poamount = CStr(CInt(drsaphead(0)))
            End If
        End If
        If (showmode = "edit") Then
            CType(CT.FindControl("txt_ponum"), TextBox).Text = po
            CType(CT.FindControl("txt_poamount"), TextBox).Text = poamount
            CType(CT.FindControl("txt_tamount"), TextBox).Text = tamount
            CType(CT.FindControl("txt_ngamount"), TextBox).Text = ngamount
        ElseIf (showmode = "showvalue") Then
            CT.Rows(2).Cells(1).Text = po
            CT.Rows(2).Cells(3).Text = poamount
            CT.Rows(2).Cells(5).Text = tamount
            CT.Rows(2).Cells(7).Text = ngamount
        End If
        If (mtype <> 0) Then
            CType(CT.FindControl("rbl_mtype"), RadioButtonList).SelectedValue = mtype
        Else
            CType(CT.FindControl("rbl_mtype"), RadioButtonList).SelectedIndex = -1
        End If
    End Sub

    Sub PutCTTail_2(showmode As String)
        Dim judge As Integer
        If (iqctype = 4) Then
            judge = drsaphead(12)
        Else
            judge = 0
        End If
        If (judge <> 0) Then
            CType(CT.FindControl("rbl_judge"), RadioButtonList).SelectedValue = judge
        Else
            CType(CT.FindControl("rbl_judge"), RadioButtonList).SelectedIndex = -1
        End If
    End Sub
    Sub PutCTTail_3(showmode As String)
        Dim inspector, auditor, cmemo, inspecdate, auditdate As String
        Dim judge As Integer
        If (iqctype = 4) Then
            If (drsaphead(15) <> "") Then
                inspector = drsaphead(15)
            Else
                inspector = "簽名"
            End If
            If (drsaphead(17) <> "") Then
                auditor = drsaphead(17)
            Else
                auditor = "簽名"
            End If
            cmemo = drsaphead(13)
            inspecdate = drsaphead(16)
            auditdate = drsaphead(18)
            judge = drsaphead(12)
        Else
            inspector = "簽名"
            auditor = "簽名"
            cmemo = ""
            inspecdate = "1900/1/1"
            auditdate = "1900/1/1"
            judge = 0
        End If
        If (showmode = "edit") Then
            CType(CT.FindControl("txt_cmemo"), TextBox).Text = cmemo
            CType(CT.FindControl("btn_inspector"), Button).Text = inspector
            CType(CT.FindControl("btn_auditor"), Button).Text = auditor
            If (judge < 1 Or auditdate <> "1900/1/1") Then
                CType(CT.FindControl("btn_inspector"), Button).Enabled = False
            Else
                CType(CT.FindControl("btn_inspector"), Button).Enabled = True
            End If
            If (inspecdate = "1900/1/1") Then
                CType(CT.FindControl("btn_auditor"), Button).Enabled = False
            Else
                CType(CT.FindControl("btn_auditor"), Button).Enabled = True
            End If
        ElseIf (showmode = "showvalue") Then
            CT.Rows(21).Cells(1).Text = cmemo
            CT.Rows(22).Cells(0).Text = inspector
            CT.Rows(22).Cells(1).Text = auditor
        End If
    End Sub
    Sub PutCTTail_4(showmode As String)
        Dim inspecdate, auditdate, amemo As String
        If (iqctype = 4) Then
            If (drsaphead(16) <> "1900/1/1") Then
                inspecdate = drsaphead(16)
            Else
                inspecdate = ""
            End If
            If (drsaphead(18) <> "1900/1/1") Then
                auditdate = drsaphead(18)
            Else
                auditdate = ""
            End If
            amemo = drsaphead(14)
        Else
            inspecdate = ""
            auditdate = ""
            amemo = ""
        End If
        If (showmode = "edit") Then
            CType(CT.FindControl("txt_amemo"), TextBox).Text = amemo
            CType(CT.FindControl("txt_inspecdate"), TextBox).Text = inspecdate
            CType(CT.FindControl("txt_auditdate"), TextBox).Text = auditdate
        ElseIf (showmode = "showvalue") Then
            CT.Rows(23).Cells(1).Text = amemo
            CT.Rows(24).Cells(0).Text = inspecdate
            CT.Rows(24).Cells(1).Text = auditdate
        End If
    End Sub
    Sub TxtResult_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim Txtx As TextBox = sender
        Dim i, j As Integer
        Dim itol, ispec, iresult As String
        i = CInt(Split(Txtx.ID, "_")(2))
        j = CInt(Split(Txtx.ID, "_")(3))
        itol = CType(CT.FindControl("txt_itol_" & i), TextBox).Text
        ispec = CType(CT.FindControl("txt_ispec_" & i), TextBox).Text
        iresult = Txtx.Text
        If (ispec <> "") Then
            If (IsNumeric(ispec)) Then
                If (IsNumeric(iresult)) Then
                    ResultJudge(i, j, itol, ispec, iresult, mode)
                Else
                    CommUtil.ShowMsg(Me, iresult & "不是數值,請check")
                End If
            End If
        Else
            Txtx.Text = ""
            CommUtil.ShowMsg(Me, "此欄位檢驗標準欄空白,不需輸入結果")
        End If
    End Sub
End Class

