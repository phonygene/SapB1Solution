**User:**

<task>
請檢查C:\Projects\SapB1Solution\MgmSP\ExpenseClaimForm.aspx.vb
根據此份檔案，推斷
C:\Projects\SapB1Solution\shopfloor\Claude_TMP\etc\roo-code-encoding-fix.md
裡的解決方案是否有被落實。

也請順便檢查worklog、DexExp等檔案，確定我設計的給AI模型使用的協作流程與架構還有規範是否完整，有沒有錯誤、矛盾。並統整一個新的檔案，報告目前AI模型協作時有被遵循的部分，並分析我的配置檔每一次Prompt會多耗用多少Token。
</task>
<environment_details>
# VSCode Visible Files
MgmSP\ExpenseClaimForm.aspx.vb

# VSCode Open Tabs
shopfloor/Claude_TMP/規格差異與實作方向確認_20251125.md,shopfloor/Claude_TMP/ExpenseClaimForm_Spec_Integration_20251126.md,DevExp/global/ui-ux.toml,agent-os/standards/global/localization.md,agent-os/config.yml,mcp-sqlserver/config.json,agent-os/standards/global/environment-health-check.md,agent-os/SESSION_INIT.md,DevExp/global/environment.toml,.git/COMMIT_EDITMSG,MgmSP/ExpenseClaimForm_DataSpec_20251203.md,MgmSP/ExpenseClaimForm.aspx,DevExp/projects/SapB1Solution/specific.toml,.editorconfig,.gitattributes,MgmSP/ExpenseClaimForm.aspx.vb,MgmSP/ExpenseClaimForm.aspx.designer.vb

# Current Time
Current time in ISO 8601 UTC format: 2025-12-04T04:02:07.582Z
User time zone: Asia/Taipei, UTC+8:00

# Current Cost
$0.00

# Current Mode
<slug>code</slug>
<name>💻 Code</name>
<model>claude-opus-4-5-20251101</model>


# Current Workspace Directory (c:/Projects/SapB1Solution) Files
_project_tree.txt
.mcp.json
compose-win.yaml
requirements.txt
SapB1Solution.sln
.claude/
.venv/
agent-os/
agent-os/config.yml
agent-os/SESSION_INIT.md
agent-os/roles/implementers.yml
agent-os/roles/verifiers.yml
collaboration-package/
collaboration-package/CONFIGURATION_GUIDE.md
collaboration-package/README.md
collaboration-package/agent-os/SESSION_INIT.md
collaboration-package/agent-os/standards/global/coding-style.md
collaboration-package/agent-os/standards/global/communication-standards.md
collaboration-package/agent-os/standards/global/localization.md
collaboration-package/agent-os/standards/global/session-management.md
collaboration-package/agent-os/standards/global/workflow-standards.md
DevExp/
DevExp/projects/SapB1Solution/specific.toml
mcp-sqlserver/
mcp-sqlserver/.claude_rules
mcp-sqlserver/.env.example
mcp-sqlserver/.env.jtdb
mcp-sqlserver/.env.JTTST
mcp-sqlserver/.gitignore
mcp-sqlserver/.python-version
mcp-sqlserver/AGENT_OS_INTEGRATION.md
mcp-sqlserver/config.json
mcp-sqlserver/DATABASE_SWITCH_GUIDE.md
mcp-sqlserver/OPERATION_RULES.md
mcp-sqlserver/pyproject.toml
mcp-sqlserver/QUICK_REFERENCE.md
mcp-sqlserver/README.md
mcp-sqlserver/SETUP_COMPLETE.md
mcp-sqlserver/switch-db.bat
mcp-sqlserver/switch-db.sh
mcp-sqlserver/test_db_direct.py
mcp-sqlserver/test_server.py
mcp-sqlserver/src/__init__.py
mcp-sqlserver/src/backup_manager.py
mcp-sqlserver/src/database.py
mcp-sqlserver/src/server.py
MgmSP/
MgmSP/Default.aspx
MgmSP/Default.aspx.designer.vb
MgmSP/Default.aspx.vb
MgmSP/ExpenseClaimForm.aspx
MgmSP/ExpenseClaimForm.aspx.bak
MgmSP/Global.asax
MgmSP/Index.aspx.designer.vb
MgmSP/Index.aspx.vb
MgmSP/kkk.aspx
MgmSP/kkk.aspx.designer.vb
MgmSP/kkk.aspx.vb
MgmSP/MgmSP.vbproj
MgmSP/MySite1.Master
MgmSP/MySite1.Master.designer.vb
MgmSP/MySite1.Master.vb
MgmSP/packages.config
MgmSP/pdfCombined.pdf
MgmSP/pdfExample.pdf
MgmSP/pdfExample1.pdf
MgmSP/rrr.aspx
MgmSP/rrr.aspx.designer.vb
MgmSP/rrr.aspx.vb
MgmSP/Web.config
MgmSP/Web.sitemap
MgmSP/App_Data/ExpenseClaimTransactionConfig.json
MgmSP/commcode/ShowData.aspx.vb
MgmSP/commcode/UserVendorTaxId.vb
MgmSP/fd/freport.aspx
MgmSP/fd/freport.aspx.designer.vb
MgmSP/fd/freport.aspx.vb
MgmSP/jFonts/kaiu.ttf
MgmSP/jFonts/msjh.ttc
MgmSP/jFonts/msjhbd.ttc
MgmSP/jFonts/msjhl.ttc
MgmSP/obj0/Debug/.NETFramework,Version=v4.0.AssemblyAttributes.vb
MgmSP/obj0/Debug/Interop.SAPbobsCOM.dll
MgmSP/obj0/Debug/MgmSP.dll
MgmSP/obj0/Debug/MgmSP.Resources.resources
MgmSP/obj0/Debug/MgmSP.vbproj.CopyComplete
MgmSP/obj0/Debug/MgmSP.vbproj.FileListAbsolute.txt
MgmSP/obj0/Debug/MgmSP.xml
MgmSP/obj0/Debug/TempPE/My Project.Application.Designer.vb.dll
MgmSP/obj0/Debug/TempPE/My Project.Resources.Designer.vb.dll
MgmSP/obj0/Release/.NETFramework,Version=v4.0.AssemblyAttributes.vb
MgmSP/obj0/Release/Interop.SAPbobsCOM.dll
MgmSP/obj0/Release/MgmSP.dll
MgmSP/obj0/Release/MgmSP.Resources.resources
MgmSP/obj0/Release/MgmSP.vbproj.FileListAbsolute.txt
MgmSP/obj0/Release/MgmSP.xml
MgmSP/obj0/Release/TempPE/My Project.Application.Designer.vb.dll
MgmSP/obj0/Release/TempPE/My Project.Resources.Designer.vb.dll
MgmSP/pu/qv.aspx
MgmSP/pu/qv.aspx.designer.vb
MgmSP/pu/qv.aspx.vb
MgmSP/signoff/agnsetup.aspx
MgmSP/signoff/agnsetup.aspx.designer.vb
MgmSP/signoff/agnsetup.aspx.vb
MgmSP/signoff/cLsignoff.aspx
MgmSP/signoff/cLsignoff.aspx.designer.vb
MgmSP/signoff/cLsignoff.aspx.vb
MgmSP/signoff/jetlog.gif
MgmSP/signoff/jetlog.jpg
MgmSP/signoff/printform.aspx
MgmSP/signoff/printform.aspx.designer.vb
MgmSP/signoff/printform.aspx.vb
MgmSP/signoff/showpdf.aspx
MgmSP/signoff/showpdf.aspx.designer.vb
MgmSP/signoff/showpdf.aspx.vb
MgmSP/signoff/signflowexclude.aspx
MgmSP/signoff/signflowexclude.aspx.designer.vb
MgmSP/signoff/signflowexclude.aspx.vb
MgmSP/signoff/signoff.aspx
MgmSP/signoff/signoff.aspx.designer.vb
MgmSP/signoff/signoff.aspx.vb
MgmSP/signoff/signoffsetup.aspx
MgmSP/signoff/signoffsetup.aspx.designer.vb
MgmSP/signoff/signoffsetup.aspx.vb
MgmSP/signoff/signofftodo.aspx
MgmSP/signoff/signofftodo.aspx.designer.vb
MgmSP/signoff/signofftodo.aspx.vb
MgmSP/signoff/signofftool.aspx
MgmSP/signoff/signofftool.aspx.designer.vb
MgmSP/signoff/signofftool.aspx.vb
MgmSP/signoff/signoffvip.aspx
MgmSP/signoff/signoffvip.aspx.vb
MgmSP/spare/sp.aspx
MgmSP/spare/sp.aspx.designer.vb
MgmSP/spare/sp.aspx.vb
MgmSP/spare/spmaterial.aspx
MgmSP/spare/spmaterial.aspx.designer.vb
MgmSP/spare/spmaterial.aspx.vb
MgmSP/spare/spold.aspx
MgmSP/spare/spold.aspx.designer.vb
MgmSP/spare/spold.aspx.vb
MgmSP/Uploads/Expense/20251126141553_Frame-74654-1.png
MgmSP/Uploads/Expense/20251126141616_paper.txt
MgmSP/usermgm/addsapid.aspx
MgmSP/usermgm/addsapid.aspx.designer.vb
MgmSP/usermgm/addsapid.aspx.vb
MgmSP/usermgm/login.aspx
MgmSP/usermgm/login.aspx.designer.vb
MgmSP/usermgm/login.aspx.vb
MgmSP/usermgm/logout.aspx
MgmSP/usermgm/logout.aspx.designer.vb
MgmSP/usermgm/logout.aspx.vb
MgmSP/usermgm/pwdchange.aspx
MgmSP/usermgm/pwdchange.aspx.designer.vb
MgmSP/usermgm/pwdchange.aspx.vb
MgmSP/usermgm/rolemodify.aspx
MgmSP/usermgm/rolemodify.aspx.designer.vb
MgmSP/usermgm/rolemodify.aspx.vb
MgmSP/usermgm/rolesetup.aspx
MgmSP/usermgm/rolesetup.aspx.designer.vb
MgmSP/usermgm/rolesetup.aspx.vb
MgmSP/usermgm/useradd.aspx
MgmSP/usermgm/useradd.aspx.designer.vb
MgmSP/usermgm/useradd.aspx.vb
MgmSP/usermgm/userlist.aspx
MgmSP/usermgm/userlist.aspx.designer.vb
MgmSP/usermgm/userlist.aspx.vb
MgmSP/usermgm/usermodify.aspx
MgmSP/usermgm/usermodify.aspx.designer.vb
MgmSP/usermgm/usermodify.aspx.vb
MgmSP/wo/issuedwoinfo.aspx
MgmSP/wo/issuedwoinfo.aspx.designer.vb
MgmSP/wo/issuedwoinfo.aspx.vb
MgmSP/wo/moadd_sys.aspx
MgmSP/wo/moadd_sys.aspx.designer.vb
MgmSP/wo/moadd_sys.aspx.vb
MgmSP/wo/molist.aspx
MgmSP/wo/molist.aspx.designer.vb
MgmSP/wo/molist.aspx.vb
MgmSP/wo/pcmaterial.aspx
MgmSP/wo/pcmaterial.aspx.designer.vb
MgmSP/wo/pcmaterial.aspx.vb
MgmSP/wo/wolist.aspx
MgmSP/wo/wolist.aspx.designer.vb
MgmSP/wo/wolist.aspx.vb
MgmSP/wo/wostamodify.aspx
MgmSP/wo/wostamodify.aspx.designer.vb
MgmSP/wo/wostamodify.aspx.vb
MgmSP/wo/wostatus.aspx
MgmSP/wo/wostatus.aspx.designer.vb
MgmSP/wo/wostatus.aspx.vb
packages/
packages/Microsoft.AspNet.WebApi.5.2.7/.signature.p7s
packages/Microsoft.AspNet.WebApi.5.2.7/Microsoft.AspNet.WebApi.5.2.7.nupkg
packages/SharpZipLib.0.86.0/.signature.p7s
shopfloor/
shopfloor/費用申請單規格TMP/實作計畫書_v1.0.md
shopfloor/費用申請單規格TMP/實作計畫書_v1.1.md
src/
worklog/

(File list truncated. Use list_files on specific subdirectories if you need to explore further.)
You have not created a todo list yet. Create one with `update_todo_list` if your task is complicated or involves multiple steps.
</environment_details>

---

**Assistant:**

我需要先讀取相關檔案來進行分析。讓我先讀取指定的檔案。

<read_file>
<args>
  <file>
    <path>MgmSP/ExpenseClaimForm.aspx.vb</path>
  </file>
  <file>
    <path>shopfloor/Claude_TMP/etc/roo-code-encoding-fix.md</path>
  </file>
</args>
</read_file>

[Response interrupted by a tool use result. Only one tool may be used at a time and should be placed at the end of the message.]

---

**User:**

[read_file for 'MgmSP/ExpenseClaimForm.aspx.vb', 'shopfloor/Claude_TMP/etc/roo-code-encoding-fix.md'] Result:
<files>
<file><path>MgmSP/ExpenseClaimForm.aspx.vb</path>
<content lines="1-1664">
   1 | ﻿Imports System.Data.SqlClient
   2 | Imports System.IO
   3 | Imports System.Web.Configuration
   4 | Imports System.Web.UI.HtmlControls
   5 | 
   6 | ''' <summary>
   7 | ''' 費用申請單 (Expense Claim Form)
   8 | ''' 重構版本: 2025-11-26 (v2.0)
   9 | ''' 對應規格: shopfloor/Claude_TMP/ExpenseClaimForm_Spec_Integration_20251126.md
  10 | ''' </summary>
  11 | Partial Public Class ExpenseClaimForm
  12 |     Inherits System.Web.UI.Page
  13 | 
  14 | #Region "類別定義"
  15 |     <Serializable()>
  16 |     Public Class ExpenseLine
  17 |         Public Property LineNum As Integer
  18 |         Public Property CategoryCode As String
  19 |         Public Property Description As String
  20 |         Public Property AcctCode As String
  21 |         Public Property LineTotal As Decimal ' 未稅金額
  22 |         Public Property VatGroup As String
  23 |         Public Property VatRate As Decimal
  24 |         Public Property VatSum As Decimal ' 稅額
  25 |         Public Property PriceAfterVat As Decimal ' 含稅金額
  26 |         Public Property CostingCode As String
  27 |         Public Property CostingCode2 As String
  28 |         ' 顯示用
  29 |         Public Property Currency As String
  30 |         Public Property Rate As Decimal
  31 |     End Class
  32 | 
  33 |     <Serializable()>
  34 |     Public Class MDRLine
  35 |         Public Property LineNum As Integer
  36 |         Public Property U_LIFNR As String
  37 |         Public Property U_STCEG As String
  38 |         Public Property U_XBLNR As String
  39 |         Public Property U_ZFORM_CODE As String
  40 |         Public Property U_BLDAT As DateTime?
  41 |         Public Property U_VATDATE As DateTime?
  42 |         Public Property U_HWBAS As Decimal ' 未稅金額
  43 |         Public Property U_HWSTE As Decimal ' 稅額
  44 |         Public Property U_TAX_TYPE As String
  45 |     End Class
  46 | 
  47 |     <Serializable()>
  48 |     Public Class AttachmentItem
  49 |         Public Property ID As Integer
  50 |         Public Property FileName As String
  51 |         Public Property FilePath As String
  52 |         Public Property UploadDate As DateTime
  53 |         Public Property UploadTime As String
  54 |         Public Property Uploader As String
  55 |     End Class
  56 | #End Region
  57 | 
  58 | #Region "變數宣告"
  59 |     Private ReadOnly connStr As String = WebConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString
  60 |     Private ReadOnly sapConnStr As String = WebConfigurationManager.ConnectionStrings("SapSQLConnection").ConnectionString
  61 | 
  62 |     Private currentUserId As String = ""
  63 |     Private currentDocEntry As Integer = 0
  64 |     Private canApprove As Boolean = False
  65 |     Private isApUser As Boolean = False ' AP_App 權限
  66 | #End Region
  67 | 
  68 | #Region "屬性 (ViewState)"
  69 |     Private Property CurrentLines As List(Of ExpenseLine)
  70 |         Get
  71 |             If ViewState("CurrentLines") Is Nothing Then
  72 |                 ViewState("CurrentLines") = New List(Of ExpenseLine)()
  73 |             End If
  74 |             Return CType(ViewState("CurrentLines"), List(Of ExpenseLine))
  75 |         End Get
  76 |         Set(value As List(Of ExpenseLine))
  77 |             ViewState("CurrentLines") = value
  78 |         End Set
  79 |     End Property
  80 | 
  81 |     Private Property CurrentMDRLines As List(Of MDRLine)
  82 |         Get
  83 |             If ViewState("CurrentMDRLines") Is Nothing Then
  84 |                 ViewState("CurrentMDRLines") = New List(Of MDRLine)()
  85 |             End If
  86 |             Return CType(ViewState("CurrentMDRLines"), List(Of MDRLine))
  87 |         End Get
  88 |         Set(value As List(Of MDRLine))
  89 |             ViewState("CurrentMDRLines") = value
  90 |         End Set
  91 |     End Property
  92 | 
  93 |     Private Property CurrentAttachments As List(Of AttachmentItem)
  94 |         Get
  95 |             If ViewState("CurrentAttachments") Is Nothing Then
  96 |                 ViewState("CurrentAttachments") = New List(Of AttachmentItem)()
  97 |             End If
  98 |             Return CType(ViewState("CurrentAttachments"), List(Of AttachmentItem))
  99 |         End Get
 100 |         Set(value As List(Of AttachmentItem))
 101 |             ViewState("CurrentAttachments") = value
 102 |         End Set
 103 |     End Property
 104 | #End Region
 105 | 
 106 | #Region "頁面載入"
 107 |     Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
 108 |         Try
 109 |             If Session("s_id") Is Nothing Then
 110 |                 Response.Redirect("~/usermgm/login.aspx")
 111 |                 Return
 112 |             End If
 113 |             currentUserId = Session("s_id").ToString()
 114 | 
 115 |             CheckApprovalPermission()
 116 | 
 117 |             If Request.QueryString("DocEntry") IsNot Nothing Then
 118 |                 Integer.TryParse(Request.QueryString("DocEntry"), currentDocEntry)
 119 |             End If
 120 | 
 121 |             If Not IsPostBack Then
 122 |                 InitializeDropDowns()
 123 | 
 124 |                 If currentDocEntry > 0 Then
 125 |                     LoadDocument(currentDocEntry)
 126 |                 Else
 127 |                     SetDefaultValues()
 128 |                     InitializeGridViews()
 129 |                 End If
 130 |             End If
 131 | 
 132 |         Catch ex As Exception
 133 |             ShowError("頁面載入錯誤: " & ex.Message)
 134 |         End Try
 135 |     End Sub
 136 | 
 137 |     Private Sub CheckApprovalPermission()
 138 |         Using conn As New SqlConnection(connStr)
 139 |             conn.Open()
 140 |             ' 檢查使用者是否有審核權限 (Approver) 與 AP作業權限 (AP_App)
 141 |             Dim sql As String = "SELECT Approver, AP_App FROM [User] WHERE id = @UserId"
 142 |             Using cmd As New SqlCommand(sql, conn)
 143 |                 cmd.Parameters.AddWithValue("@UserId", currentUserId)
 144 |                 Using dr As SqlDataReader = cmd.ExecuteReader()
 145 |                     If dr.Read() Then
 146 |                         canApprove = (Convert.ToInt32(dr("Approver")) = 1)
 147 |                         isApUser = (Convert.ToInt32(dr("AP_App")) = 1)
 148 |                     End If
 149 |                 End Using
 150 |             End Using
 151 |         End Using
 152 | 
 153 |         ' 預設不隱藏，由 LoadDocument 或 Page_Load 決定顯示狀態
 154 |     End Sub
 155 | 
 156 |     Private Sub SetDefaultValues()
 157 |         lblDocNum.Text = "[新單據]"
 158 |         txtOwner.Text = currentUserId
 159 |         If ddlPurchaser.Items.Count > 0 Then ddlPurchaser.SelectedValue = currentUserId
 160 |         If ddlPurchaser.SelectedIndex = -1 AndAlso ddlPurchaser.Items.Count > 0 Then ddlPurchaser.SelectedIndex = 0
 161 |         lblDocStatus.Text = "草稿"
 162 |         lblDocStatus.CssClass = "badge status-P"
 163 |         txtStatusDisplay.Text = "草稿"
 164 | 
 165 |         Dim today As String = DateTime.Now.ToString("yyyy-MM-dd")
 166 |         txtDocDate.Text = today
 167 |         txtTaxDate.Text = today
 168 |         txtDocDueDate.Text = DateTime.Now.AddDays(30).ToString("yyyy-MM-dd")
 169 | 
 170 |         If ddlDocCurrency.Items.FindByValue("NTD") IsNot Nothing Then
 171 |             ddlDocCurrency.SelectedValue = "NTD"
 172 |         ElseIf ddlDocCurrency.Items.FindByValue("TWD") IsNot Nothing Then
 173 |             ddlDocCurrency.SelectedValue = "TWD"
 174 |         End If
 175 |         txtDocRate.Text = "1.0"
 176 | 
 177 |         ' 新增單據時，審核區塊預設唯讀且停用按鈕 (但保持顯示以確認版面)
 178 |         txtApprovalComments.ReadOnly = True
 179 |         
 180 |         btnApprove.Visible = True
 181 |         btnApprove.Enabled = False
 182 |         
 183 |         btnUpdateComment.Visible = True
 184 |         btnUpdateComment.Enabled = False
 185 |         
 186 |         btnReject.Visible = True
 187 |         btnReject.Enabled = False
 188 |     End Sub
 189 | #End Region
 190 | 
 191 | #Region "初始化資料"
 192 |     Private Sub InitializeDropDowns()
 193 |         LoadDeliveryAddress()
 194 |         LoadCurrencies()
 195 |         LoadPaymentGroups()
 196 |         LoadPurchasers()
 197 |     End Sub
 198 | 
 199 |     Private Sub InitializeGridViews()
 200 |         If CurrentLines.Count = 0 Then
 201 |             ' 預設加入一筆空行
 202 |             AddNewEmptyLine()
 203 |         End If
 204 |         BindGrid()
 205 | 
 206 |         If CurrentMDRLines.Count = 0 Then
 207 |             ' MDR 預設不加空行，由使用者手動新增
 208 |         End If
 209 |         BindMDRGrid()
 210 |         
 211 |         BindAttachmentGrid()
 212 |     End Sub
 213 | 
 214 |     Private Sub LoadDeliveryAddress()
 215 |         ddlDeliveryAddr.Items.Clear()
 216 |         ddlDeliveryAddr.Items.Add(New ListItem("- 請選擇 -", ""))
 217 |         Try
 218 |             Using conn As New SqlConnection(connStr)
 219 |                 conn.Open()
 220 |                 Dim sql As String = "SELECT ID, addrName, address FROM addr WHERE addrType='R' AND active='Y'"
 221 |                 Using cmd As New SqlCommand(sql, conn)
 222 |                     Using dr As SqlDataReader = cmd.ExecuteReader()
 223 |                         While dr.Read()
 224 |                             Dim item As New ListItem(dr("addrName").ToString(), dr("ID").ToString())
 225 |                             item.Attributes.Add("data-address", dr("address").ToString())
 226 |                             ddlDeliveryAddr.Items.Add(item)
 227 |                         End While
 228 |                     End Using
 229 |                 End Using
 230 |             End Using
 231 |         Catch ex As Exception
 232 |             ShowError("載入地址失敗: " & ex.Message)
 233 |         End Try
 234 |     End Sub
 235 | 
 236 |     Private Sub LoadCurrencies()
 237 |         ddlDocCurrency.Items.Clear()
 238 |         Try
 239 |             Using conn As New SqlConnection(sapConnStr)
 240 |                 conn.Open()
 241 |                 Dim sql As String = "SELECT CurrCode FROM OCRN"
 242 |                 Using cmd As New SqlCommand(sql, conn)
 243 |                     Using dr As SqlDataReader = cmd.ExecuteReader()
 244 |                         While dr.Read()
 245 |                             ddlDocCurrency.Items.Add(New ListItem(dr("CurrCode").ToString(), dr("CurrCode").ToString()))
 246 |                         End While
 247 |                     End Using
 248 |                 End Using
 249 |             End Using
 250 |         Catch ex As Exception
 251 |             ShowError("載入幣別失敗: " & ex.Message)
 252 |         End Try
 253 |     End Sub
 254 | 
 255 |     Private Sub LoadPaymentGroups()
 256 |         ddlGroupNum.Items.Clear()
 257 |         Try
 258 |             Using conn As New SqlConnection(sapConnStr)
 259 |                 conn.Open()
 260 |                 Dim sql As String = "SELECT GroupNum, PymntGroup FROM OCTG"
 261 |                 Using cmd As New SqlCommand(sql, conn)
 262 |                     Using dr As SqlDataReader = cmd.ExecuteReader()
 263 |                         While dr.Read()
 264 |                             ddlGroupNum.Items.Add(New ListItem(dr("PymntGroup").ToString(), dr("GroupNum").ToString()))
 265 |                         End While
 266 |                     End Using
 267 |                 End Using
 268 |             End Using
 269 |         Catch ex As Exception
 270 |             ShowError("載入付款條件失敗: " & ex.Message)
 271 |         End Try
 272 |     End Sub
 273 | 
 274 |     Private Sub LoadPurchasers()
 275 |         ddlPurchaser.Items.Clear()
 276 |         Try
 277 |             Using conn As New SqlConnection(sapConnStr)
 278 |                 conn.Open()
 279 |                 Dim sql As String = "SELECT SlpName FROM OSLP"
 280 |                 Using cmd As New SqlCommand(sql, conn)
 281 |                     Using dr As SqlDataReader = cmd.ExecuteReader()
 282 |                         While dr.Read()
 283 |                             ddlPurchaser.Items.Add(New ListItem(dr("SlpName").ToString(), dr("SlpName").ToString()))
 284 |                         End While
 285 |                     End Using
 286 |                 End Using
 287 |             End Using
 288 |         Catch ex As Exception
 289 |             ShowError("載入採購人員失敗: " & ex.Message)
 290 |         End Try
 291 |     End Sub
 292 | 
 293 |     ' 在 GridView RowDataBound 時呼叫
 294 |     Private Sub LoadExpenseCategories(ddl As DropDownList)
 295 |         ddl.Items.Clear()
 296 |         ddl.Items.Add(New ListItem("-選擇-", ""))
 297 |         Try
 298 |             Using conn As New SqlConnection(connStr)
 299 |                 conn.Open()
 300 |                 Dim sql As String = "SELECT CategoryCode, CategoryName, AcctCode FROM expense_category WHERE Active='Y'"
 301 |                 Using cmd As New SqlCommand(sql, conn)
 302 |                     Using dr As SqlDataReader = cmd.ExecuteReader()
 303 |                         While dr.Read()
 304 |                             Dim item As New ListItem(dr("CategoryName").ToString(), dr("CategoryCode").ToString())
 305 |                             item.Attributes.Add("data-acct", dr("AcctCode").ToString())
 306 |                             ddl.Items.Add(item)
 307 |                         End While
 308 |                     End Using
 309 |                 End Using
 310 |             End Using
 311 |         Catch ex As Exception
 312 |             ShowError("載入費用類別失敗")
 313 |         End Try
 314 |     End Sub
 315 | 
 316 |     Private Sub LoadVatGroups(ddl As DropDownList)
 317 |         ddl.Items.Clear()
 318 |         ' 1-應稅 (5%), 2-零稅 (0), 3-免稅 (0)
 319 |         Dim item1 As New ListItem("1-應稅 (5%)", "1")
 320 |         item1.Attributes.Add("data-rate", "5")
 321 |         ddl.Items.Add(item1)
 322 | 
 323 |         Dim item2 As New ListItem("2-零稅 (0%)", "2")
 324 |         item2.Attributes.Add("data-rate", "0")
 325 |         ddl.Items.Add(item2)
 326 | 
 327 |         Dim item3 As New ListItem("3-免稅 (0%)", "3")
 328 |         item3.Attributes.Add("data-rate", "0")
 329 |         ddl.Items.Add(item3)
 330 |     End Sub
 331 | 
 332 |     Private Sub LoadProducts(ddl As DropDownList)
 333 |         ddl.Items.Clear()
 334 |         ddl.Items.Add(New ListItem("", ""))
 335 |         Try
 336 |             Using conn As New SqlConnection(sapConnStr)
 337 |                 conn.Open()
 338 |                 Dim sql As String = "SELECT DimCode, DimDesc FROM ODIM"
 339 |                 Using cmd As New SqlCommand(sql, conn)
 340 |                     Using dr As SqlDataReader = cmd.ExecuteReader()
 341 |                         While dr.Read()
 342 |                             ddl.Items.Add(New ListItem(dr("DimDesc").ToString(), dr("DimCode").ToString()))
 343 |                         End While
 344 |                     End Using
 345 |                 End Using
 346 |             End Using
 347 |         Catch ex As Exception
 348 |             ShowError("載入產品失敗")
 349 |         End Try
 350 |     End Sub
 351 | 
 352 |     Private Sub LoadDepartments(ddl As DropDownList)
 353 |         ddl.Items.Clear()
 354 |         ddl.Items.Add(New ListItem("", ""))
 355 |         Try
 356 |             Using conn As New SqlConnection(sapConnStr)
 357 |                 conn.Open()
 358 |                 Dim sql As String = "SELECT PrcCode, PrcName FROM OPRC"
 359 |                 Using cmd As New SqlCommand(sql, conn)
 360 |                     Using dr As SqlDataReader = cmd.ExecuteReader()
 361 |                         While dr.Read()
 362 |                             ddl.Items.Add(New ListItem(dr("PrcName").ToString() & " (" & dr("PrcCode").ToString() & ")", dr("PrcCode").ToString()))
 363 |                         End While
 364 |                     End Using
 365 |                 End Using
 366 |             End Using
 367 |         Catch ex As Exception
 368 |             ShowError("載入部門失敗")
 369 |         End Try
 370 |     End Sub
 371 | #End Region
 372 | 
 373 | #Region "供應商搜尋 (Search Modal)"
 374 |     Protected Sub btnSearchCardCode_Click(sender As Object, e As EventArgs)
 375 |         PerformVendorSearch("Code", txtCardCode.Text.Trim())
 376 |     End Sub
 377 | 
 378 |     Protected Sub btnSearchCardName_Click(sender As Object, e As EventArgs)
 379 |         PerformVendorSearch("Name", txtCardName.Text.Trim())
 380 |     End Sub
 381 | 
 382 |     Private Sub PerformVendorSearch(source As String, keyword As String)
 383 |         hfSearchSource.Value = source
 384 |         txtVendorSearchKeyword.Text = keyword
 385 |         BindVendorSearchGrid(keyword)
 386 |         mpeVendor.Show()
 387 |         pnlVendorSearch.Style("display") = "block"
 388 |     End Sub
 389 | 
 390 |     Protected Sub btnDoSearchVendor_Click(sender As Object, e As EventArgs)
 391 |         BindVendorSearchGrid(txtVendorSearchKeyword.Text.Trim())
 392 |         mpeVendor.Show()
 393 |     End Sub
 394 | 
 395 |     Protected Sub gvVendorSearch_PageIndexChanging(sender As Object, e As GridViewPageEventArgs)
 396 |         gvVendorSearch.PageIndex = e.NewPageIndex
 397 |         BindVendorSearchGrid(txtVendorSearchKeyword.Text.Trim())
 398 |         mpeVendor.Show()
 399 |     End Sub
 400 | 
 401 |     Private Sub BindVendorSearchGrid(keyword As String)
 402 |         Try
 403 |             Dim sqlWhere As String = "WHERE CardType='S' AND FrozenFor='N' "
 404 | 
 405 |             Dim searchSource As String = hfSearchSource.Value ' 判斷搜尋來源
 406 |             Dim isExact As Boolean = (rblSearchMode.SelectedValue = "Exact")
 407 | 
 408 |             If Not String.IsNullOrEmpty(keyword) Then
 409 |                 keyword = keyword.Replace("*", "").Replace("%", "")
 410 | 
 411 |                 If isExact Then
 412 |                     ' 開頭比對
 413 |                     If searchSource = "Code" Then
 414 |                         sqlWhere &= " AND CardCode LIKE @Kw"
 415 |                     Else ' Name
 416 |                         sqlWhere &= " AND CardName LIKE @Kw"
 417 |                     End If
 418 |                     keyword = keyword & "%"
 419 |                 Else
 420 |                     ' 模糊比對
 421 |                     If searchSource = "Code" Then
 422 |                         sqlWhere &= " AND CardCode LIKE @Kw"
 423 |                     Else ' Name
 424 |                         sqlWhere &= " AND CardName LIKE @Kw"
 425 |                     End If
 426 |                     keyword = "%" & keyword & "%"
 427 |                 End If
 428 |             End If
 429 | 
 430 |             Using conn As New SqlConnection(sapConnStr)
 431 |                 conn.Open()
 432 |                 Dim sql As String = $"SELECT CardCode, CardName FROM OCRD {sqlWhere} ORDER BY CardCode"
 433 |                 Using cmd As New SqlCommand(sql, conn)
 434 |                     If Not String.IsNullOrEmpty(keyword) Then
 435 |                         cmd.Parameters.AddWithValue("@Kw", keyword)
 436 |                     End If
 437 | 
 438 |                     Using da As New SqlDataAdapter(cmd)
 439 |                         Dim dt As New DataTable()
 440 |                         da.Fill(dt)
 441 |                         gvVendorSearch.DataSource = dt
 442 |                         gvVendorSearch.DataBind()
 443 |                     End Using
 444 |                 End Using
 445 |             End Using
 446 |         Catch ex As Exception
 447 |             ShowError("搜尋供應商錯誤: " & ex.Message)
 448 |         End Try
 449 |     End Sub
 450 | 
 451 |     Protected Sub gvVendorSearch_RowCommand(sender As Object, e As GridViewCommandEventArgs)
 452 |         If e.CommandName = "SelectVendor" Then
 453 |             Dim args As String() = e.CommandArgument.ToString().Split("|"c)
 454 |             If args.Length >= 2 Then
 455 |                 txtCardCode.Text = args(0)
 456 |                 txtCardName.Text = args(1)
 457 | 
 458 |                 ' 連動資料 (幣別, 付款條件)
 459 |                 LoadVendorLinkedData(args(0))
 460 | 
 461 |                 ' 清除錯誤訊息
 462 |                 lblErrCardCode.Visible = False
 463 |                 lblErrCardName.Visible = False
 464 |             End If
 465 |             mpeVendor.Hide()
 466 |             pnlVendorSearch.Style("display") = "none"
 467 |         End If
 468 |     End Sub
 469 | 
 470 |     Private Sub LoadVendorLinkedData(cardCode As String)
 471 |         Try
 472 |             Using conn As New SqlConnection(sapConnStr)
 473 |                 conn.Open()
 474 |                 ' 幣別 (包含 U_LastCur 檢查)
 475 |                 Dim sqlCurr As String = "SELECT Currency, U_LastCur FROM OCRD WHERE CardCode = @CardCode"
 476 |                 Dim currencyAll As Boolean = False
 477 |                 Using cmd As New SqlCommand(sqlCurr, conn)
 478 |                     cmd.Parameters.AddWithValue("@CardCode", cardCode)
 479 |                     Using dr As SqlDataReader = cmd.ExecuteReader()
 480 |                         If dr.Read() Then
 481 |                             Dim dbCurr As String = dr("Currency").ToString()
 482 |                             Dim lastCur As String = If(IsDBNull(dr("U_LastCur")), "", dr("U_LastCur").ToString())
 483 | 
 484 |                             If dbCurr = "##" Then
 485 |                                 currencyAll = True
 486 |                                 ddlDocCurrency.Enabled = True
 487 |                                 ' 優先帶入 U_LastCur，若無則帶入 NTD 或 TWD
 488 |                                 If Not String.IsNullOrEmpty(lastCur) AndAlso ddlDocCurrency.Items.FindByValue(lastCur) IsNot Nothing Then
 489 |                                     ddlDocCurrency.SelectedValue = lastCur
 490 |                                 Else
 491 |                                     If ddlDocCurrency.Items.FindByValue("NTD") IsNot Nothing Then
 492 |                                         ddlDocCurrency.SelectedValue = "NTD"
 493 |                                     ElseIf ddlDocCurrency.Items.FindByValue("TWD") IsNot Nothing Then
 494 |                                         ddlDocCurrency.SelectedValue = "TWD"
 495 |                                     End If
 496 |                                 End If
 497 |                             Else
 498 |                                 currencyAll = False
 499 |                                 ddlDocCurrency.Enabled = False ' 鎖定幣別
 500 |                                 If ddlDocCurrency.Items.FindByValue(dbCurr) IsNot Nothing Then
 501 |                                     ddlDocCurrency.SelectedValue = dbCurr
 502 |                                 End If
 503 |                             End If
 504 |                         End If
 505 |                     End Using
 506 |                 End Using
 507 | 
 508 |                 ' 觸發匯率更新與明細連動
 509 |                 ddlDocCurrency_SelectedIndexChanged(Nothing, Nothing)
 510 | 
 511 |                 ' 付款條件
 512 |                 Dim sqlPymnt As String = "SELECT GroupNum FROM OCRD WHERE CardCode = @CardCode"
 513 |                 Using cmd As New SqlCommand(sqlPymnt, conn)
 514 |                     cmd.Parameters.AddWithValue("@CardCode", cardCode)
 515 |                     Dim grp As Object = cmd.ExecuteScalar()
 516 |                     If grp IsNot Nothing AndAlso Not IsDBNull(grp) Then
 517 |                         Dim grpNum As String = grp.ToString()
 518 |                         If ddlGroupNum.Items.FindByValue(grpNum) IsNot Nothing Then
 519 |                             ddlGroupNum.SelectedValue = grpNum
 520 |                         End If
 521 |                     End If
 522 |                 End Using
 523 |             End Using
 524 |         Catch ex As Exception
 525 |             ShowError("載入供應商關聯資料失敗: " & ex.Message)
 526 |         End Try
 527 |     End Sub
 528 | #End Region
 529 | 
 530 | #Region "表頭欄位連動"
 531 |     Protected Sub ddlDeliveryAddr_SelectedIndexChanged(sender As Object, e As EventArgs)
 532 |         Dim id As String = ddlDeliveryAddr.SelectedValue
 533 |         If String.IsNullOrEmpty(id) Then
 534 |             txtAddress.Text = ""
 535 |             Return
 536 |         End If
 537 | 
 538 |         Using conn As New SqlConnection(connStr)
 539 |             conn.Open()
 540 |             Dim sql As String = "SELECT address FROM addr WHERE ID=@ID"
 541 |             Using cmd As New SqlCommand(sql, conn)
 542 |                 cmd.Parameters.AddWithValue("@ID", id)
 543 |                 Dim res = cmd.ExecuteScalar()
 544 |                 If res IsNot Nothing Then
 545 |                     txtAddress.Text = res.ToString()
 546 |                 End If
 547 |             End Using
 548 |         End Using
 549 |     End Sub
 550 | 
 551 |     Protected Sub ddlDocCurrency_SelectedIndexChanged(sender As Object, e As EventArgs)
 552 |         Dim curr As String = ddlDocCurrency.SelectedValue
 553 |         Dim rate As Decimal = 1D
 554 | 
 555 |         If curr = "TWD" Then
 556 |             txtDocRate.Text = "1.0"
 557 |         Else
 558 |             ' 取得 DocDate，若無值則預設為今天
 559 |             Dim docDate As DateTime = DateTime.Today
 560 |             If Not String.IsNullOrEmpty(txtDocDate.Text) Then
 561 |                 DateTime.TryParse(txtDocDate.Text, docDate)
 562 |             End If
 563 | 
 564 |             Using conn As New SqlConnection(sapConnStr)
 565 |                 conn.Open()
 566 |                 Dim sql As String = "SELECT TOP 1 Rate FROM ORTT WHERE Currency=@Curr AND RateDate <= @DocDate ORDER BY RateDate DESC"
 567 |                 Using cmd As New SqlCommand(sql, conn)
 568 |                     cmd.Parameters.AddWithValue("@Curr", curr)
 569 |                     cmd.Parameters.AddWithValue("@DocDate", docDate)
 570 |                     Dim res = cmd.ExecuteScalar()
 571 |                     If res IsNot Nothing Then
 572 |                         rate = Convert.ToDecimal(res)
 573 |                         txtDocRate.Text = rate.ToString("F4")
 574 |                     Else
 575 |                         txtDocRate.Text = "1.0"
 576 |                     End If
 577 |                 End Using
 578 |             End Using
 579 |         End If
 580 | 
 581 |         ' 同步更新明細列的幣別與匯率
 582 |         UpdateLinesCurrency(curr, rate)
 583 |     End Sub
 584 | 
 585 |     Protected Sub txtDocRate_TextChanged(sender As Object, e As EventArgs)
 586 |         Dim curr As String = ddlDocCurrency.SelectedValue
 587 |         Dim rate As Decimal = 1D
 588 |         Decimal.TryParse(txtDocRate.Text, rate)
 589 | 
 590 |         ' 同步更新明細列的幣別與匯率
 591 |         UpdateLinesCurrency(curr, rate)
 592 |     End Sub
 593 | 
 594 |     Private Sub UpdateLinesCurrency(curr As String, rate As Decimal)
 595 |         Dim lines = CurrentLines
 596 |         For Each line In lines
 597 |             line.Currency = curr
 598 |             line.Rate = rate
 599 |         Next
 600 |         CurrentLines = lines
 601 |         BindGrid()
 602 |     End Sub
 603 | 
 604 | 
 605 |     Protected Sub btnRefreshRate_Click(sender As Object, e As EventArgs)
 606 |         ddlDocCurrency_SelectedIndexChanged(Nothing, Nothing)
 607 |     End Sub
 608 | #End Region
 609 | 
 610 | #Region "Expense GridView"
 611 |     Protected Sub btnAddLine_Click(sender As Object, e As EventArgs)
 612 |         SyncGridDataToModel()
 613 |         AddNewEmptyLine()
 614 |         BindGrid()
 615 |     End Sub
 616 | 
 617 |     Private Sub AddNewEmptyLine()
 618 |         Dim lines = CurrentLines
 619 |         Dim newNum As Integer = 1
 620 |         If lines.Count > 0 Then newNum = lines.Max(Function(x) x.LineNum) + 1
 621 | 
 622 |         Dim curr As String = ddlDocCurrency.SelectedValue
 623 |         Dim rate As Decimal = 1D
 624 |         Decimal.TryParse(txtDocRate.Text, rate)
 625 | 
 626 |         lines.Add(New ExpenseLine With {
 627 |             .LineNum = newNum,
 628 |             .LineTotal = 0,
 629 |             .VatRate = 0,
 630 |             .PriceAfterVat = 0,
 631 |             .Currency = curr,
 632 |             .Rate = rate
 633 |         })
 634 |         CurrentLines = lines
 635 |     End Sub
 636 | 
 637 |     Protected Sub btnDeleteLine_Click(sender As Object, e As EventArgs)
 638 |         SyncGridDataToModel()
 639 |         Dim lines = CurrentLines
 640 |         Dim newLines As New List(Of ExpenseLine)
 641 | 
 642 |         For Each row As GridViewRow In gvExpenseDetail.Rows
 643 |             If row.RowType = DataControlRowType.DataRow Then
 644 |                 Dim chk As CheckBox = CType(row.FindControl("chkSelect"), CheckBox)
 645 |                 If chk IsNot Nothing AndAlso Not chk.Checked Then
 646 |                     Dim idx As Integer = row.DataItemIndex
 647 |                     If idx < lines.Count Then newLines.Add(lines(idx))
 648 |                 End If
 649 |             End If
 650 |         Next
 651 | 
 652 |         ' 重算 LineNum
 653 |         For i As Integer = 0 To newLines.Count - 1
 654 |             newLines(i).LineNum = i + 1
 655 |         Next
 656 | 
 657 |         CurrentLines = newLines
 658 |         BindGrid()
 659 |     End Sub
 660 | 
 661 |     Private Sub BindGrid()
 662 |         gvExpenseDetail.DataSource = CurrentLines
 663 |         gvExpenseDetail.DataBind()
 664 |         CalculateFooterTotal()
 665 |     End Sub
 666 | 
 667 |     Protected Sub gvExpenseDetail_RowDataBound(sender As Object, e As GridViewRowEventArgs)
 668 |         If e.Row.RowType = DataControlRowType.DataRow Then
 669 |             Dim line As ExpenseLine = CType(e.Row.DataItem, ExpenseLine)
 670 | 
 671 |             ' Expense Category
 672 |             Dim ddlCat As DropDownList = CType(e.Row.FindControl("ddlExpCategory"), DropDownList)
 673 |             LoadExpenseCategories(ddlCat)
 674 |             If ddlCat.Items.FindByValue(line.CategoryCode) IsNot Nothing Then ddlCat.SelectedValue = line.CategoryCode
 675 | 
 676 |             ' Vat Group
 677 |             Dim ddlVat As DropDownList = CType(e.Row.FindControl("ddlVatGroup"), DropDownList)
 678 |             LoadVatGroups(ddlVat)
 679 |             If ddlVat.Items.FindByValue(line.VatGroup) IsNot Nothing Then ddlVat.SelectedValue = line.VatGroup
 680 | 
 681 |             ' Products (CostingCode)
 682 |             Dim ddlCost As DropDownList = CType(e.Row.FindControl("ddlCostingCode"), DropDownList)
 683 |             LoadProducts(ddlCost)
 684 |             If ddlCost.Items.FindByValue(line.CostingCode) IsNot Nothing Then ddlCost.SelectedValue = line.CostingCode
 685 | 
 686 |             ' Departments (CostingCode2)
 687 |             Dim ddlCost2 As DropDownList = CType(e.Row.FindControl("ddlCostingCode2"), DropDownList)
 688 |             LoadDepartments(ddlCost2)
 689 |             If ddlCost2.Items.FindByValue(line.CostingCode2) IsNot Nothing Then ddlCost2.SelectedValue = line.CostingCode2
 690 | 
 691 |             ' Values
 692 |             CType(e.Row.FindControl("txtDescription"), TextBox).Text = line.Description
 693 |             CType(e.Row.FindControl("txtAcctCode"), TextBox).Text = line.AcctCode
 694 |             CType(e.Row.FindControl("txtLineTotal"), TextBox).Text = line.LineTotal.ToString("0.##")
 695 |             CType(e.Row.FindControl("txtVatSum"), TextBox).Text = line.VatSum.ToString("0.##")
 696 |             CType(e.Row.FindControl("txtPriceAfterVat"), TextBox).Text = line.PriceAfterVat.ToString("0.##")
 697 | 
 698 |             ' Currency & Rate
 699 |             CType(e.Row.FindControl("txtLineCurrency"), TextBox).Text = line.Currency
 700 |             CType(e.Row.FindControl("txtLineRate"), TextBox).Text = line.Rate.ToString("0.####")
 701 |         End If
 702 |     End Sub
 703 | 
 704 |     Protected Sub gvExpenseDetail_RowCommand(sender As Object, e As GridViewCommandEventArgs)
 705 |         ' 保留未來擴充
 706 |     End Sub
 707 | 
 708 |     Protected Sub ddlExpCategory_SelectedIndexChanged(sender As Object, e As EventArgs)
 709 |         Dim ddl As DropDownList = CType(sender, DropDownList)
 710 |         Dim row As GridViewRow = CType(ddl.NamingContainer, GridViewRow)
 711 |         Dim txtAcct As TextBox = CType(row.FindControl("txtAcctCode"), TextBox)
 712 | 
 713 |         If ddl.SelectedIndex > 0 Then
 714 |             txtAcct.Text = ddl.SelectedItem.Attributes("data-acct")
 715 |         Else
 716 |             txtAcct.Text = ""
 717 |         End If
 718 | 
 719 |         SyncGridDataToModel()
 720 |     End Sub
 721 | 
 722 |     Protected Sub txtDescription_TextChanged(sender As Object, e As EventArgs)
 723 |         SyncGridDataToModel()
 724 |     End Sub
 725 | 
 726 |     Protected Sub CalculateLineTotal(sender As Object, e As EventArgs)
 727 |         ' 當未稅金額、稅別改變時觸發
 728 |         SyncGridDataToModel(False)
 729 | 
 730 |         Dim ddl As DropDownList = TryCast(sender, DropDownList)
 731 |         Dim txt As TextBox = TryCast(sender, TextBox)
 732 |         Dim row As GridViewRow = Nothing
 733 |         If ddl IsNot Nothing Then row = CType(ddl.NamingContainer, GridViewRow)
 734 |         If txt IsNot Nothing Then row = CType(txt.NamingContainer, GridViewRow)
 735 | 
 736 |         If row IsNot Nothing Then
 737 |             Dim rowIndex As Integer = row.DataItemIndex
 738 |             Dim lines = CurrentLines
 739 |             If rowIndex < lines.Count Then
 740 |                 Dim line = lines(rowIndex)
 741 |                 ' 重新計算稅額與含稅金額
 742 |                 If line.VatGroup = "1" Then ' 1-應稅 (5%)
 743 |                     line.VatRate = 5
 744 |                     line.VatSum = Math.Round(line.LineTotal * 0.05D, 0)
 745 |                 Else
 746 |                     line.VatRate = 0
 747 |                     line.VatSum = 0
 748 |                 End If
 749 |                 line.PriceAfterVat = line.LineTotal + line.VatSum
 750 |             End If
 751 |             CurrentLines = lines
 752 |         End If
 753 | 
 754 |         BindGrid()
 755 |     End Sub
 756 | 
 757 |     Protected Sub CalculateVatSum(sender As Object, e As EventArgs)
 758 |         ' 當稅額被手動修改時
 759 |         SyncGridDataToModel(True) ' Update Model with screen VatSum
 760 | 
 761 |         Dim lines = CurrentLines
 762 |         Dim txt As TextBox = CType(sender, TextBox)
 763 |         Dim row As GridViewRow = CType(txt.NamingContainer, GridViewRow)
 764 |         Dim rowIndex As Integer = row.DataItemIndex
 765 | 
 766 |         If rowIndex < lines.Count Then
 767 |             Dim line = lines(rowIndex)
 768 |             ' 手動修改稅額後，重新計算含稅金額
 769 |             line.PriceAfterVat = line.LineTotal + line.VatSum
 770 |         End If
 771 |         CurrentLines = lines
 772 |         BindGrid()
 773 |     End Sub
 774 | 
 775 |     Protected Sub CalculatePriceAfterVat(sender As Object, e As EventArgs)
 776 |         ' 當含稅金額改變時觸發 -> 反算未稅金額
 777 |         Dim txt As TextBox = CType(sender, TextBox)
 778 |         Dim row As GridViewRow = CType(txt.NamingContainer, GridViewRow)
 779 |         Dim rowIndex As Integer = row.DataItemIndex
 780 | 
 781 |         Dim priceAfterVat As Decimal = 0
 782 |         Decimal.TryParse(txt.Text, priceAfterVat)
 783 | 
 784 |         SyncGridDataToModel(False)
 785 | 
 786 |         Dim lines = CurrentLines
 787 |         If rowIndex < lines.Count Then
 788 |             Dim line = lines(rowIndex)
 789 |             line.PriceAfterVat = priceAfterVat
 790 | 
 791 |             ' 反算邏輯
 792 |             ' 1-應稅 (5%)
 793 |             If line.VatGroup = "1" Then
 794 |                 line.VatRate = 5
 795 |                 line.LineTotal = Math.Round(priceAfterVat / 1.05D, 0)
 796 |                 line.VatSum = priceAfterVat - line.LineTotal
 797 |             Else
 798 |                 line.VatRate = 0
 799 |                 line.LineTotal = priceAfterVat
 800 |                 line.VatSum = 0
 801 |             End If
 802 |         End If
 803 |         CurrentLines = lines
 804 |         BindGrid()
 805 |     End Sub
 806 | 
 807 |     Private Sub SyncGridDataToModel(Optional readPriceAfterVat As Boolean = False)
 808 |         Dim lines = CurrentLines
 809 |         For i As Integer = 0 To gvExpenseDetail.Rows.Count - 1
 810 |             Dim row As GridViewRow = gvExpenseDetail.Rows(i)
 811 |             If i < lines.Count Then
 812 |                 Dim ddlExp As DropDownList = CType(row.FindControl("ddlExpCategory"), DropDownList)
 813 |                 Dim txtDesc As TextBox = CType(row.FindControl("txtDescription"), TextBox)
 814 |                 Dim txtAcct As TextBox = CType(row.FindControl("txtAcctCode"), TextBox)
 815 |                 Dim txtTotal As TextBox = CType(row.FindControl("txtLineTotal"), TextBox)
 816 |                 Dim ddlVat As DropDownList = CType(row.FindControl("ddlVatGroup"), DropDownList)
 817 |                 Dim txtVatSum As TextBox = CType(row.FindControl("txtVatSum"), TextBox)
 818 |                 Dim txtPrice As TextBox = CType(row.FindControl("txtPriceAfterVat"), TextBox)
 819 |                 Dim ddlCost As DropDownList = CType(row.FindControl("ddlCostingCode"), DropDownList)
 820 |                 Dim ddlCost2 As DropDownList = CType(row.FindControl("ddlCostingCode2"), DropDownList)
 821 | 
 822 |                 Dim line = lines(i)
 823 |                 If ddlExp IsNot Nothing Then line.CategoryCode = ddlExp.SelectedValue
 824 |                 If txtDesc IsNot Nothing Then line.Description = txtDesc.Text
 825 |                 If txtAcct IsNot Nothing Then line.AcctCode = txtAcct.Text
 826 |                 If txtTotal IsNot Nothing Then Decimal.TryParse(txtTotal.Text, line.LineTotal)
 827 |                 If ddlVat IsNot Nothing Then line.VatGroup = ddlVat.SelectedValue
 828 |                 If txtVatSum IsNot Nothing Then Decimal.TryParse(txtVatSum.Text, line.VatSum)
 829 |                 If ddlCost IsNot Nothing Then line.CostingCode = ddlCost.SelectedValue
 830 |                 If ddlCost2 IsNot Nothing Then line.CostingCode2 = ddlCost2.SelectedValue
 831 | 
 832 |                 ' 注意：這裡不再自動重算 VatSum，避免覆蓋使用者手動輸入的稅額。
 833 |                 ' 重算邏輯移至 CalculateLineTotal (當 LineTotal 或 VatGroup 改變時)
 834 |                 ' 或 CalculatePriceAfterVat (當 PriceAfterVat 改變時)
 835 | 
 836 |                 ' 若是由 CalculatePriceAfterVat 呼叫 (readPriceAfterVat=False)，這時 UI 上的 PriceAfterVat 可能還沒寫入 Model (因為還沒 Assign)，
 837 |                 ' 但這裡是讀取 UI 上的 VatSum。
 838 |                 ' 基本上 SyncGridDataToModel 負責將畫面上的值(包含使用者手動改的)同步回 Model。
 839 |             End If
 840 |         Next
 841 |         CurrentLines = lines
 842 |     End Sub
 843 | #End Region
 844 | 
 845 | #Region "MDR GridView"
 846 |     Protected Sub btnAddMDRRow_Click(sender As Object, e As EventArgs)
 847 |         hfActiveTab.Value = "mdr" ' 保持在 MDR 頁籤
 848 |         SyncMDRGridToModel()
 849 |         Dim lines = CurrentMDRLines
 850 |         Dim newNum As Integer = 1
 851 |         If lines.Count > 0 Then newNum = lines.Max(Function(x) x.LineNum) + 1
 852 | 
 853 |         lines.Add(New MDRLine With {
 854 |             .LineNum = newNum,
 855 |             .U_LIFNR = txtCardCode.Text, ' 預設帶入表頭供應商
 856 |             .U_BLDAT = DateTime.Now,
 857 |             .U_VATDATE = DateTime.Now,
 858 |             .U_HWBAS = 0,
 859 |             .U_HWSTE = 0,
 860 |             .U_TAX_TYPE = "1", ' 預設應稅
 861 |             .U_ZFORM_CODE = "21" ' 預設發票類型
 862 |         })
 863 |         CurrentMDRLines = lines
 864 |         BindMDRGrid()
 865 |     End Sub
 866 | 
 867 |     Protected Sub btnDeleteMDRRow_Click(sender As Object, e As EventArgs)
 868 |         hfActiveTab.Value = "mdr" ' 保持在 MDR 頁籤
 869 |         SyncMDRGridToModel()
 870 |         Dim lines = CurrentMDRLines
 871 |         Dim newLines As New List(Of MDRLine)
 872 | 
 873 |         For Each row As GridViewRow In gvMDRDetail.Rows
 874 |             If row.RowType = DataControlRowType.DataRow Then
 875 |                 Dim chk As CheckBox = CType(row.FindControl("chkSelectMDR"), CheckBox)
 876 |                 If chk IsNot Nothing AndAlso Not chk.Checked Then
 877 |                     Dim idx As Integer = row.DataItemIndex
 878 |                     If idx < lines.Count Then newLines.Add(lines(idx))
 879 |                 End If
 880 |             End If
 881 |         Next
 882 | 
 883 |         ' 重算 LineNum
 884 |         For i As Integer = 0 To newLines.Count - 1
 885 |             newLines(i).LineNum = i + 1
 886 |         Next
 887 | 
 888 |         CurrentMDRLines = newLines
 889 |         BindMDRGrid()
 890 |     End Sub
 891 | 
 892 |     Private Sub BindMDRGrid()
 893 |         gvMDRDetail.DataSource = CurrentMDRLines
 894 |         gvMDRDetail.DataBind()
 895 |     End Sub
 896 | 
 897 |     Protected Sub gvMDRDetail_RowDataBound(sender As Object, e As GridViewRowEventArgs)
 898 |         ' 簡單綁定，大部分是 TextBox 直接在前端顯示
 899 |     End Sub
 900 | 
 901 |     Protected Sub CalculateMDRTotal(sender As Object, e As EventArgs)
 902 |         SyncMDRGridToModel(False)
 903 |         BindMDRGrid()
 904 |     End Sub
 905 | 
 906 |     Protected Sub CalculateMDRTaxManual(sender As Object, e As EventArgs)
 907 |         SyncMDRGridToModel(True) ' 手動修改稅額
 908 |         BindMDRGrid()
 909 |     End Sub
 910 | 
 911 |     Private Sub SyncMDRGridToModel(Optional isManualTax As Boolean = False)
 912 |         Dim lines = CurrentMDRLines
 913 |         For i As Integer = 0 To gvMDRDetail.Rows.Count - 1
 914 |             Dim row As GridViewRow = gvMDRDetail.Rows(i)
 915 |             If i < lines.Count Then
 916 |                 Dim line = lines(i)
 917 | 
 918 |                 Dim txtLIFNR As TextBox = CType(row.FindControl("txtLIFNR"), TextBox)
 919 |                 Dim txtSTCEG As TextBox = CType(row.FindControl("txtSTCEG"), TextBox)
 920 |                 Dim txtXBLNR As TextBox = CType(row.FindControl("txtXBLNR"), TextBox)
 921 |                 Dim ddlZFORM As DropDownList = CType(row.FindControl("ddlZFORM_CODE"), DropDownList)
 922 |                 Dim txtBLDAT As TextBox = CType(row.FindControl("txtBLDAT"), TextBox)
 923 |                 Dim txtVATDATE As TextBox = CType(row.FindControl("txtVATDATE"), TextBox)
 924 |                 Dim txtHWBAS As TextBox = CType(row.FindControl("txtHWBAS"), TextBox)
 925 |                 Dim txtHWSTE As TextBox = CType(row.FindControl("txtHWSTE"), TextBox)
 926 |                 Dim ddlTAX As DropDownList = CType(row.FindControl("ddlTAX_TYPE"), DropDownList)
 927 | 
 928 |                 If txtLIFNR IsNot Nothing Then line.U_LIFNR = txtLIFNR.Text
 929 |                 If txtSTCEG IsNot Nothing Then line.U_STCEG = txtSTCEG.Text
 930 |                 If txtXBLNR IsNot Nothing Then line.U_XBLNR = txtXBLNR.Text
 931 |                 If ddlZFORM IsNot Nothing Then line.U_ZFORM_CODE = ddlZFORM.SelectedValue
 932 | 
 933 |                 Dim dt As DateTime
 934 |                 If txtBLDAT IsNot Nothing AndAlso DateTime.TryParse(txtBLDAT.Text, dt) Then line.U_BLDAT = dt
 935 |                 If txtVATDATE IsNot Nothing AndAlso DateTime.TryParse(txtVATDATE.Text, dt) Then line.U_VATDATE = dt
 936 | 
 937 |                 If txtHWBAS IsNot Nothing Then Decimal.TryParse(txtHWBAS.Text, line.U_HWBAS)
 938 |                 If ddlTAX IsNot Nothing Then line.U_TAX_TYPE = ddlTAX.SelectedValue
 939 |                 
 940 |                 ' 若手動修改，則讀取 UI 上的值
 941 |                 If isManualTax AndAlso txtHWSTE IsNot Nothing Then
 942 |                     Decimal.TryParse(txtHWSTE.Text, line.U_HWSTE)
 943 |                 End If
 944 | 
 945 |                 ' 自動計算稅額 (僅在非手動模式下)
 946 |                 If Not isManualTax Then
 947 |                     ' 1-應稅 (5%), 2-零稅 (0), 3-免稅 (0)
 948 |                     If line.U_TAX_TYPE = "1" Then
 949 |                         line.U_HWSTE = Math.Round(line.U_HWBAS * 0.05D, 0)
 950 |                     Else
 951 |                         line.U_HWSTE = 0
 952 |                     End If
 953 |                 End If
 954 |             End If
 955 |         Next
 956 |         CurrentMDRLines = lines
 957 |     End Sub
 958 | 
 959 |     Protected Sub txtXBLNR_TextChanged(sender As Object, e As EventArgs)
 960 |         ' 當發票號碼改變時，自動判斷類型
 961 |         Dim txt As TextBox = CType(sender, TextBox)
 962 |         Dim row As GridViewRow = CType(txt.NamingContainer, GridViewRow)
 963 |         Dim ddlZForm As DropDownList = CType(row.FindControl("ddlZFORM_CODE"), DropDownList)
 964 | 
 965 |         Dim invNum As String = txt.Text.Trim()
 966 |         Dim formCode As String = "99" ' Default: 其他
 967 | 
 968 |         ' 判斷邏輯
 969 |         ' 1. 統一發票: 2碼英文 + 8碼數字
 970 |         ' 2. 海關報關單: 3碼英文
 971 |         ' 3. 高鐵票: 數字
 972 |         ' 4. 公營事業: BB或BBN
 973 |         ' 5. 其他: 警告
 974 | 
 975 |         If System.Text.RegularExpressions.Regex.IsMatch(invNum, "^[A-Z]{2}\d{8}$") Then
 976 |             formCode = "21"
 977 |         ElseIf System.Text.RegularExpressions.Regex.IsMatch(invNum, "^[A-Z]{3}") Then
 978 |             formCode = "28"
 979 |         ElseIf System.Text.RegularExpressions.Regex.IsMatch(invNum, "^(BB|BBN)") Then
 980 |             formCode = "22"
 981 |         ElseIf System.Text.RegularExpressions.Regex.IsMatch(invNum, "^\d+$") Then
 982 |             formCode = "24"
 983 |         Else
 984 |             formCode = "99"
 985 |         End If
 986 | 
 987 |         If ddlZForm IsNot Nothing Then
 988 |             If ddlZForm.Items.FindByValue(formCode) IsNot Nothing Then
 989 |                 ddlZForm.SelectedValue = formCode
 990 |             End If
 991 |         End If
 992 | 
 993 |         SyncMDRGridToModel()
 994 |     End Sub
 995 | #End Region
 996 | 
 997 | #Region "總計與表尾"
 998 |     Private Sub CalculateFooterTotal()
 999 |         Dim docTotal As Decimal = CurrentLines.Sum(Function(x) x.LineTotal)
1000 |         Dim vatSum As Decimal = CurrentLines.Sum(Function(x) x.VatSum)
1001 |         Dim gTotal As Decimal = CurrentLines.Sum(Function(x) x.PriceAfterVat)
1002 | 
1003 |         lblDocTotal.Text = docTotal.ToString("N2")
1004 |         lblVatSum.Text = vatSum.ToString("N2")
1005 |         lblDocTotalWithTax.Text = gTotal.ToString("N2")
1006 |     End Sub
1007 | #End Region
1008 | 
1009 | #Region "存檔與讀取"
1010 |     Protected Sub btnSave_Click(sender As Object, e As EventArgs)
1011 |         If Not ValidateAll() Then Return
1012 |         SaveDocument("P") ' P = 草稿
1013 |     End Sub
1014 | 
1015 |     Protected Sub btnSubmit_Click(sender As Object, e As EventArgs)
1016 |         If Not ValidateAll() Then Return
1017 |         SaveDocument("W") ' W = 待審核
1018 |     End Sub
1019 | 
1020 | 
1021 | 
1022 |     Private Property WarningConfirmed As Boolean
1023 |         Get
1024 |             If ViewState("WarningConfirmed") Is Nothing Then Return False
1025 |             Return Convert.ToBoolean(ViewState("WarningConfirmed"))
1026 |         End Get
1027 |         Set(value As Boolean)
1028 |             ViewState("WarningConfirmed") = value
1029 |         End Set
1030 |     End Property
1031 | 
1032 |     Private Function ValidateAll() As Boolean
1033 |         Dim isValid As Boolean = True
1034 |         lblMessage.Text = ""
1035 | 
1036 |         If String.IsNullOrEmpty(txtCardCode.Text) Then
1037 |             lblErrCardCode.Text = "必填"
1038 |             lblErrCardCode.Visible = True
1039 |             isValid = False
1040 |         End If
1041 | 
1042 |         If String.IsNullOrEmpty(txtDocDate.Text) Then
1043 |             lblErrDocDate.Text = "必填"
1044 |             lblErrDocDate.Visible = True
1045 |             isValid = False
1046 |         End If
1047 | 
1048 |         If String.IsNullOrEmpty(txtDocDueDate.Text) Then
1049 |             lblErrDocDueDate.Text = "必填"
1050 |             lblErrDocDueDate.Visible = True
1051 |             isValid = False
1052 |         End If
1053 | 
1054 |         ' 明細檢查
1055 |         If CurrentLines.Count = 0 AndAlso CurrentMDRLines.Count = 0 Then
1056 |             ShowError("請至少新增一筆費用明細或發票明細")
1057 |             isValid = False
1058 |         End If
1059 | 
1060 |         ' 檢查是否有 "99-其他" 類型的發票且未確認
1061 |         Dim hasOtherType As Boolean = CurrentMDRLines.Any(Function(x) x.U_ZFORM_CODE = "99")
1062 |         If hasOtherType AndAlso Not WarningConfirmed Then
1063 |             ShowError("警告：有「其他」類型的單據，請確認格式是否正確。若確定要新增，請再次點擊儲存/送出。")
1064 |             lblMessage.ForeColor = System.Drawing.Color.Orange
1065 |             WarningConfirmed = True ' 設定已確認旗標
1066 |             isValid = False
1067 |         End If
1068 | 
1069 |         Return isValid
1070 |     End Function
1071 | 
1072 |     Private Function SaveDocument(status As String, Optional isAutoSave As Boolean = False) As Boolean
1073 |         Try
1074 |             SyncGridDataToModel()
1075 |             SyncMDRGridToModel()
1076 | 
1077 |             Using conn As New SqlConnection(connStr)
1078 |                 conn.Open()
1079 |                 Using trans = conn.BeginTransaction()
1080 |                     Try
1081 |                         Dim jID As Integer = 0
1082 | 
1083 |                         ' 1. jOPCH (Header)
1084 |                         If currentDocEntry = 0 Then
1085 |                             ' Insert
1086 |                             Dim sqlH As String = "INSERT INTO jOPCH (CardCode, CardName, NumAtCard, InvNum, DeliveryAddrID, AddressName, Address, " &
1087 |                                                "DocDate, DocDueDate, TaxDate, DocCurrency, DocRate, DocTotal, VatSum, " &
1088 |                                                "GroupNum, Comments, ApprovalStatus, CreateBy, CreateDate, U_PID) " &
1089 |                                                "VALUES (@CardCode, @CardName, @NumAtCard, @InvNum, @DeliveryAddrID, @AddressName, @Address, " &
1090 |                                                "@DocDate, @DocDueDate, @TaxDate, @DocCurrency, @DocRate, @DocTotal, @VatSum, " &
1091 |                                                "@GroupNum, @Comments, @Status, @User, GETDATE(), @UPID); " &
1092 |                                                "SELECT SCOPE_IDENTITY();"
1093 | 
1094 |                             Using cmd As New SqlCommand(sqlH, conn, trans)
1095 |                                 SetHeaderParameters(cmd, status)
1096 |                                 jID = Convert.ToInt32(cmd.ExecuteScalar())
1097 |                             End Using
1098 | 
1099 |                             ' Update DocEntry = jID
1100 |                             Dim sqlUpd As String = "UPDATE jOPCH SET DocEntry=@ID, DocNum=@ID WHERE jID=@ID"
1101 |                             Using cmd As New SqlCommand(sqlUpd, conn, trans)
1102 |                                 cmd.Parameters.AddWithValue("@ID", jID)
1103 |                                 cmd.ExecuteNonQuery()
1104 |                             End Using
1105 |                             currentDocEntry = jID
1106 |                         Else
1107 |                             ' Update
1108 |                             jID = currentDocEntry
1109 |                             Dim sqlH As String = "UPDATE jOPCH SET CardCode=@CardCode, CardName=@CardName, NumAtCard=@NumAtCard, InvNum=@InvNum, " &
1110 |                                                "DeliveryAddrID=@DeliveryAddrID, AddressName=@AddressName, Address=@Address, " &
1111 |                                                "DocDate=@DocDate, DocDueDate=@DocDueDate, TaxDate=@TaxDate, DocCurrency=@DocCurrency, " &
1112 |                                                "DocRate=@DocRate, DocTotal=@DocTotal, VatSum=@VatSum, " &
1113 |                                                "GroupNum=@GroupNum, Comments=@Comments, ApprovalStatus=@Status, " &
1114 |                                                "UpdateBy=@User, UpdateDate=GETDATE(), U_PID=@UPID WHERE DocEntry=@ID"
1115 | 
1116 |                             Using cmd As New SqlCommand(sqlH, conn, trans)
1117 |                                 cmd.Parameters.AddWithValue("@ID", currentDocEntry)
1118 |                                 SetHeaderParameters(cmd, status)
1119 |                                 cmd.ExecuteNonQuery()
1120 |                             End Using
1121 | 
1122 |                             ' Delete old lines
1123 |                             Using cmd As New SqlCommand("DELETE FROM jPCH1 WHERE DocEntry=@ID", conn, trans)
1124 |                                 cmd.Parameters.AddWithValue("@ID", currentDocEntry)
1125 |                                 cmd.ExecuteNonQuery()
1126 |                             End Using
1127 | 
1128 |                             ' Delete old MDR
1129 |                             Using cmd As New SqlCommand("DELETE FROM jMGUIAP WHERE DocEntry=@ID", conn, trans)
1130 |                                 cmd.Parameters.AddWithValue("@ID", currentDocEntry)
1131 |                                 cmd.ExecuteNonQuery()
1132 |                             End Using
1133 |                             Using cmd As New SqlCommand("DELETE FROM jMGUIAPDetail WHERE DocEntry=@ID", conn, trans)
1134 |                                 cmd.Parameters.AddWithValue("@ID", currentDocEntry)
1135 |                                 cmd.ExecuteNonQuery()
1136 |                             End Using
1137 |                         End If
1138 | 
1139 |                         ' 2. jPCH1 (Expense Lines)
1140 |                         Dim sqlL As String = "INSERT INTO jPCH1 (jID, DocEntry, LineNum, ItemCode, Dscription, AcctCode, " &
1141 |                                            "LineTotal, VatGroup, VatPrcnt, LineVat, GTotal, CostingCode, CostingCode2) " &
1142 |                                            "VALUES (@jID, @DocEntry, @LineNum, @ItemCode, @Dscription, @AcctCode, " &
1143 |                                            "@LineTotal, @VatGroup, @VatPrcnt, @LineVat, @GTotal, @CostingCode, @CostingCode2)"
1144 | 
1145 |                         For Each line As ExpenseLine In CurrentLines
1146 |                             Using cmd As New SqlCommand(sqlL, conn, trans)
1147 |                                 cmd.Parameters.AddWithValue("@jID", jID) ' FK to Header jID
1148 |                                 cmd.Parameters.AddWithValue("@DocEntry", currentDocEntry)
1149 |                                 cmd.Parameters.AddWithValue("@LineNum", line.LineNum)
1150 |                                 cmd.Parameters.AddWithValue("@ItemCode", line.CategoryCode)
1151 |                                 cmd.Parameters.AddWithValue("@Dscription", line.Description)
1152 |                                 cmd.Parameters.AddWithValue("@AcctCode", line.AcctCode)
1153 |                                 cmd.Parameters.AddWithValue("@LineTotal", line.LineTotal)
1154 |                                 cmd.Parameters.AddWithValue("@VatGroup", line.VatGroup)
1155 |                                 cmd.Parameters.AddWithValue("@VatPrcnt", line.VatRate)
1156 |                                 cmd.Parameters.AddWithValue("@LineVat", line.VatSum)
1157 |                                 cmd.Parameters.AddWithValue("@GTotal", line.PriceAfterVat)
1158 |                                 cmd.Parameters.AddWithValue("@CostingCode", line.CostingCode)
1159 |                                 cmd.Parameters.AddWithValue("@CostingCode2", line.CostingCode2)
1160 |                                 cmd.ExecuteNonQuery()
1161 |                             End Using
1162 |                         Next
1163 | 
1164 |                         ' 3. jMGUIAP & jMGUIAPDetail (MDR)
1165 |                         If CurrentMDRLines.Count > 0 Then
1166 |                             ' MDR Header (彙總)
1167 |                             Dim mdrTotal As Decimal = CurrentMDRLines.Sum(Function(x) x.U_HWBAS)
1168 |                             Dim mdrVat As Decimal = CurrentMDRLines.Sum(Function(x) x.U_HWSTE)
1169 | 
1170 |                             Dim sqlMdrH As String = "INSERT INTO jMGUIAP (DocEntry, DocNum, DocTotal, VatSum, CreateBy, CreateDate) " &
1171 |                                                   "VALUES (@DocEntry, @DocNum, @DocTotal, @VatSum, @User, GETDATE()); SELECT SCOPE_IDENTITY();"
1172 | 
1173 |                             Dim mdrID As Integer = 0
1174 |                             Using cmd As New SqlCommand(sqlMdrH, conn, trans)
1175 |                                 cmd.Parameters.AddWithValue("@DocEntry", currentDocEntry)
1176 |                                 cmd.Parameters.AddWithValue("@DocNum", currentDocEntry)
1177 |                                 cmd.Parameters.AddWithValue("@DocTotal", mdrTotal)
1178 |                                 cmd.Parameters.AddWithValue("@VatSum", mdrVat)
1179 |                                 cmd.Parameters.AddWithValue("@User", currentUserId)
1180 |                                 mdrID = Convert.ToInt32(cmd.ExecuteScalar())
1181 |                             End Using
1182 | 
1183 |                             ' MDR Lines
1184 |                             Dim sqlMdrL As String = "INSERT INTO jMGUIAPDetail (jID, DocEntry, LineNum, U_LIFNR, U_STCEG, U_XBLNR, U_ZFORM_CODE, " &
1185 |                                                   "U_BLDAT, U_VATDATE, U_HWBAS, U_HWSTE, U_TAX_TYPE) " &
1186 |                                                   "VALUES (@jID, @DocEntry, @LineNum, @LIFNR, @STCEG, @XBLNR, @ZFORM, @BLDAT, @VATDATE, @HWBAS, @HWSTE, @TAXTYPE)"
1187 | 
1188 |                             For Each line As MDRLine In CurrentMDRLines
1189 |                                 Using cmd As New SqlCommand(sqlMdrL, conn, trans)
1190 |                                     cmd.Parameters.AddWithValue("@jID", mdrID)
1191 |                                     cmd.Parameters.AddWithValue("@DocEntry", currentDocEntry)
1192 |                                     cmd.Parameters.AddWithValue("@LineNum", line.LineNum)
1193 |                                     cmd.Parameters.AddWithValue("@LIFNR", line.U_LIFNR)
1194 |                                     cmd.Parameters.AddWithValue("@STCEG", line.U_STCEG)
1195 |                                     cmd.Parameters.AddWithValue("@XBLNR", line.U_XBLNR)
1196 |                                     cmd.Parameters.AddWithValue("@ZFORM", line.U_ZFORM_CODE)
1197 |                                     cmd.Parameters.AddWithValue("@BLDAT", If(line.U_BLDAT.HasValue, line.U_BLDAT.Value, DBNull.Value))
1198 |                                     cmd.Parameters.AddWithValue("@VATDATE", If(line.U_VATDATE.HasValue, line.U_VATDATE.Value, DBNull.Value))
1199 |                                     cmd.Parameters.AddWithValue("@HWBAS", line.U_HWBAS)
1200 |                                     cmd.Parameters.AddWithValue("@HWSTE", line.U_HWSTE)
1201 |                                     cmd.Parameters.AddWithValue("@TAXTYPE", line.U_TAX_TYPE)
1202 |                                     cmd.ExecuteNonQuery()
1203 |                                 End Using
1204 |                             Next
1205 |                         End If
1206 | 
1207 |                         ' Update U_LastCur in OCRD if needed
1208 |                         ' (若表頭幣別為多幣別的業務夥伴，每次新增單據時更新 U_LastCur)
1209 |                         ' 這裡簡化邏輯：直接嘗試更新，若 U_LastCur 不存在會報錯，但我們假設已建立
1210 |                         If ddlDocCurrency.Enabled Then ' 表示該業務夥伴為多幣別 (##)
1211 |                             Dim sqlUpdCur As String = "UPDATE OCRD SET U_LastCur = @Curr WHERE CardCode = @CardCode"
1212 |                             Using cmd As New SqlCommand(sqlUpdCur, conn, trans)
1213 |                                 ' 注意：OCRD 在 SAP DB，這裡 conn 是 jtdb...
1214 |                                 ' 跨庫交易較複雜，我們可能需要另開連線到 SAP DB，但不參與此 Transaction (風險可接受)
1215 |                             End Using
1216 | 
1217 |                             ' 分開執行 SAP DB Update
1218 |                             Using sapConn As New SqlConnection(sapConnStr)
1219 |                                 sapConn.Open()
1220 |                                 Using cmdSap As New SqlCommand(sqlUpdCur, sapConn)
1221 |                                     cmdSap.Parameters.AddWithValue("@Curr", ddlDocCurrency.SelectedValue)
1222 |                                     cmdSap.Parameters.AddWithValue("@CardCode", txtCardCode.Text)
1223 |                                     cmdSap.ExecuteNonQuery()
1224 |                                 End Using
1225 |                             End Using
1226 |                         End If
1227 | 
1228 |                         trans.Commit()
1229 |                         lblMessage.Text = "儲存成功"
1230 |                         lblDocNum.Text = currentDocEntry.ToString()
1231 | 
1232 |                         If Not isAutoSave Then
1233 |                             Response.Redirect("ExpenseClaimForm.aspx?DocEntry=" & currentDocEntry)
1234 |                         End If
1235 |                         Return True
1236 |                     Catch ex As Exception
1237 |                         trans.Rollback()
1238 |                         Throw ex
1239 |                     End Try
1240 |                 End Using
1241 |             End Using
1242 |         Catch ex As Exception
1243 |             ShowError("儲存失敗: " & ex.Message)
1244 |             Return False
1245 |         End Try
1246 |     End Function
1247 |     Protected Sub btnDelete_Click(sender As Object, e As EventArgs)
1248 |         If currentDocEntry > 0 Then
1249 |             Try
1250 |                 Using conn As New SqlConnection(connStr)
1251 |                     conn.Open()
1252 |                     Using trans = conn.BeginTransaction()
1253 |                         Try
1254 |                             ' Delete Lines
1255 |                             Dim cmd As New SqlCommand("DELETE FROM jPCH1 WHERE DocEntry=@ID", conn, trans)
1256 |                             cmd.Parameters.AddWithValue("@ID", currentDocEntry)
1257 |                             cmd.ExecuteNonQuery()
1258 | 
1259 |                             ' Delete MDR
1260 |                             cmd.CommandText = "DELETE FROM jMGUIAP WHERE DocEntry=@ID"
1261 |                             cmd.ExecuteNonQuery()
1262 |                             cmd.CommandText = "DELETE FROM jMGUIAPDetail WHERE DocEntry=@ID"
1263 |                             cmd.ExecuteNonQuery()
1264 | 
1265 |                             ' Delete Header
1266 |                             cmd.CommandText = "DELETE FROM jOPCH WHERE DocEntry=@ID"
1267 |                             cmd.ExecuteNonQuery()
1268 |                             
1269 |                             ' Delete Attachments
1270 |                             cmd.CommandText = "DELETE FROM jAttach WHERE jID=@ID"
1271 |                             cmd.ExecuteNonQuery()
1272 | 
1273 |                             trans.Commit()
1274 |                             Response.Redirect("Index.aspx")
1275 |                         Catch ex As Exception
1276 |                             trans.Rollback()
1277 |                             Throw ex
1278 |                         End Try
1279 |                     End Using
1280 |                 End Using
1281 |             Catch ex As Exception
1282 |                 ShowError("刪除失敗: " & ex.Message)
1283 |             End Try
1284 |         End If
1285 |     End Sub
1286 | 
1287 |     Protected Sub btnCancel_Click(sender As Object, e As EventArgs)
1288 |         Response.Redirect("Index.aspx")
1289 |     End Sub
1290 | 
1291 |     Private Sub SetHeaderParameters(cmd As SqlCommand, status As String)
1292 |         cmd.Parameters.AddWithValue("@CardCode", txtCardCode.Text)
1293 |         cmd.Parameters.AddWithValue("@CardName", txtCardName.Text)
1294 |         cmd.Parameters.AddWithValue("@NumAtCard", txtNumAtCard.Text)
1295 |         'cmd.Parameters.AddWithValue("@InvNum", txtInvNum.Text)
1296 |         cmd.Parameters.AddWithValue("@DeliveryAddrID", ddlDeliveryAddr.SelectedValue)
1297 |         cmd.Parameters.AddWithValue("@AddressName", ddlDeliveryAddr.SelectedItem.Text)
1298 |         cmd.Parameters.AddWithValue("@Address", txtAddress.Text)
1299 |         cmd.Parameters.AddWithValue("@DocDate", txtDocDate.Text)
1300 |         cmd.Parameters.AddWithValue("@DocDueDate", txtDocDueDate.Text)
1301 |         cmd.Parameters.AddWithValue("@TaxDate", If(String.IsNullOrEmpty(txtTaxDate.Text), DBNull.Value, txtTaxDate.Text))
1302 |         cmd.Parameters.AddWithValue("@DocCurrency", ddlDocCurrency.SelectedValue)
1303 |         cmd.Parameters.AddWithValue("@DocRate", txtDocRate.Text)
1304 |         cmd.Parameters.AddWithValue("@DocTotal", Decimal.Parse(lblDocTotalWithTax.Text.Replace(",", "")))
1305 |         cmd.Parameters.AddWithValue("@VatSum", Decimal.Parse(lblVatSum.Text.Replace(",", "")))
1306 |         cmd.Parameters.AddWithValue("@GroupNum", ddlGroupNum.SelectedValue)
1307 |         cmd.Parameters.AddWithValue("@Comments", txtRemarks.Text)
1308 |         'cmd.Parameters.AddWithValue("@Attachment", If(String.IsNullOrEmpty(lblAttachment.Text), DBNull.Value, lblAttachment.Text))
1309 |         cmd.Parameters.AddWithValue("@Status", status)
1310 |         cmd.Parameters.AddWithValue("@User", currentUserId)
1311 |         cmd.Parameters.AddWithValue("@UPID", If(String.IsNullOrEmpty(txtUPID.Text), DBNull.Value, txtUPID.Text))
1312 |     End Sub
1313 | #End Region
1314 | 
1315 | #Region "讀取文件"
1316 |     Private Sub LoadDocument(id As Integer)
1317 |         Using conn As New SqlConnection(connStr)
1318 |             conn.Open()
1319 |             ' Load Header
1320 |             Dim sql As String = "SELECT * FROM jOPCH WHERE DocEntry=@ID"
1321 |             Using cmd As New SqlCommand(sql, conn)
1322 |                 cmd.Parameters.AddWithValue("@ID", id)
1323 |                 Using dr As SqlDataReader = cmd.ExecuteReader()
1324 |                     If dr.Read() Then
1325 |                         lblDocNum.Text = dr("DocEntry").ToString()
1326 |                         txtJID.Text = dr("jID").ToString()
1327 | 
1328 |                         txtCardCode.Text = dr("CardCode").ToString()
1329 |                         txtCardName.Text = dr("CardName").ToString()
1330 |                         txtNumAtCard.Text = dr("NumAtCard").ToString()
1331 |                         'txtInvNum.Text = dr("InvNum").ToString()
1332 | 
1333 |                         If Not IsDBNull(dr("DeliveryAddrID")) Then ddlDeliveryAddr.SelectedValue = dr("DeliveryAddrID").ToString()
1334 |                         txtAddress.Text = dr("Address").ToString()
1335 | 
1336 |                         txtDocDate.Text = Convert.ToDateTime(dr("DocDate")).ToString("yyyy-MM-dd")
1337 |                         If Not IsDBNull(dr("DocDueDate")) Then txtDocDueDate.Text = Convert.ToDateTime(dr("DocDueDate")).ToString("yyyy-MM-dd")
1338 |                         If Not IsDBNull(dr("TaxDate")) Then txtTaxDate.Text = Convert.ToDateTime(dr("TaxDate")).ToString("yyyy-MM-dd")
1339 | 
1340 |                         ddlDocCurrency.SelectedValue = dr("DocCurrency").ToString()
1341 |                         txtDocRate.Text = dr("DocRate").ToString()
1342 | 
1343 |                         If Not IsDBNull(dr("GroupNum")) Then ddlGroupNum.SelectedValue = dr("GroupNum").ToString()
1344 |                         txtRemarks.Text = dr("Comments").ToString()
1345 | 
1346 |                         'If Not IsDBNull(dr("Attachment")) Then
1347 |                         '    lblAttachment.Text = dr("Attachment").ToString()
1348 |                         'End If
1349 | 
1350 |                         Dim status As String = dr("ApprovalStatus").ToString()
1351 |                         lblDocStatus.Text = GetStatusText(status)
1352 |                         lblDocStatus.CssClass = "badge status-" & status
1353 |                         txtStatusDisplay.Text = GetStatusText(status)
1354 |                         txtApprovalStatus.Text = status
1355 | 
1356 |                         txtApprovedBy.Text = dr("ApprovedBy").ToString()
1357 |                         txtOwner.Text = dr("CreateBy").ToString()
1358 |                         If ddlPurchaser.Items.FindByValue(dr("CreateBy").ToString()) IsNot Nothing Then
1359 |                             ddlPurchaser.SelectedValue = dr("CreateBy").ToString()
1360 |                         End If
1361 | 
1362 |                         If Not IsDBNull(dr("U_PID")) Then txtUPID.Text = dr("U_PID").ToString()
1363 | 
1364 |                         ' 顯示審核區塊邏輯:
1365 |                         ' 1. 這個區塊要 jtdb 的 User Table 裡的 AP_App 欄位為 1 才可以編輯
1366 |                         ' 2. 一般 User 不能編輯也不能按放行/發送意見/退回
1367 |                         ' 3. 但為了查看退回意見等，建議顯示但唯讀
1368 | 
1369 |                         pnlApproval.Visible = True ' 只要是既有單據，預設顯示，內容依權限控制
1370 | 
1371 |                         txtApprovalComments.ReadOnly = Not isApUser
1372 |                         
1373 |                         ' 按鈕保持顯示，但無權限者停用
1374 |                         btnApprove.Visible = True
1375 |                         btnApprove.Enabled = isApUser
1376 |                         
1377 |                         btnUpdateComment.Visible = True
1378 |                         btnUpdateComment.Enabled = isApUser
1379 |                         
1380 |                         btnReject.Visible = True
1381 |                         btnReject.Enabled = isApUser
1382 |                     End If
1383 |                 End Using
1384 |             End Using
1385 | 
1386 |             ' Load Attachments (jAttach)
1387 |             Dim attachList As New List(Of AttachmentItem)
1388 |             sql = "SELECT * FROM jAttach WHERE jID=@ID ORDER BY UploadDate, UploadTime"
1389 |             Using cmd As New SqlCommand(sql, conn)
1390 |                 cmd.Parameters.AddWithValue("@ID", id)
1391 |                 Using dr As SqlDataReader = cmd.ExecuteReader()
1392 |                     While dr.Read()
1393 |                         attachList.Add(New AttachmentItem With {
1394 |                             .ID = Convert.ToInt32(dr("ID")),
1395 |                             .FileName = dr("FilePath").ToString(), ' 這裡暫存 FileName (存檔名)
1396 |                             .FilePath = dr("FilePath").ToString(),
1397 |                             .UploadDate = Convert.ToDateTime(dr("UploadDate")),
1398 |                             .UploadTime = dr("UploadTime").ToString(),
1399 |                             .Uploader = dr("Uploader").ToString()
1400 |                         })
1401 |                     End While
1402 |                 End Using
1403 |             End Using
1404 |             CurrentAttachments = attachList
1405 |             BindAttachmentGrid()
1406 | 
1407 |             ' Load Expense Lines
1408 |             Dim lines As New List(Of ExpenseLine)
1409 |             sql = "SELECT * FROM jPCH1 WHERE DocEntry=@ID ORDER BY LineNum"
1410 |             Using cmd As New SqlCommand(sql, conn)
1411 |                 cmd.Parameters.AddWithValue("@ID", id)
1412 |                 Using dr As SqlDataReader = cmd.ExecuteReader()
1413 |                     While dr.Read()
1414 |                         lines.Add(New ExpenseLine With {
1415 |                             .LineNum = Convert.ToInt32(dr("LineNum")),
1416 |                             .CategoryCode = dr("ItemCode").ToString(),
1417 |                             .Description = dr("Dscription").ToString(),
1418 |                             .AcctCode = dr("AcctCode").ToString(),
1419 |                             .LineTotal = Convert.ToDecimal(dr("LineTotal")),
1420 |                             .VatGroup = dr("VatGroup").ToString(),
1421 |                             .VatRate = Convert.ToDecimal(dr("VatPrcnt")),
1422 |                             .VatSum = Convert.ToDecimal(dr("LineVat")),
1423 |                             .PriceAfterVat = Convert.ToDecimal(dr("GTotal")),
1424 |                             .CostingCode = dr("CostingCode").ToString(),
1425 |                             .CostingCode2 = dr("CostingCode2").ToString()
1426 |                         })
1427 |                     End While
1428 |                 End Using
1429 |             End Using
1430 |             CurrentLines = lines
1431 |             BindGrid()
1432 | 
1433 |             ' Load MDR Lines
1434 |             Dim mdrLines As New List(Of MDRLine)
1435 |             sql = "SELECT * FROM jMGUIAPDetail WHERE DocEntry=@ID ORDER BY LineNum"
1436 |             Using cmd As New SqlCommand(sql, conn)
1437 |                 cmd.Parameters.AddWithValue("@ID", id)
1438 |                 Using dr As SqlDataReader = cmd.ExecuteReader()
1439 |                     While dr.Read()
1440 |                         Dim mdr As New MDRLine With {
1441 |                             .LineNum = Convert.ToInt32(dr("LineNum")),
1442 |                             .U_LIFNR = dr("U_LIFNR").ToString(),
1443 |                             .U_STCEG = dr("U_STCEG").ToString(),
1444 |                             .U_XBLNR = dr("U_XBLNR").ToString(),
1445 |                             .U_ZFORM_CODE = dr("U_ZFORM_CODE").ToString(),
1446 |                             .U_HWBAS = Convert.ToDecimal(dr("U_HWBAS")),
1447 |                             .U_HWSTE = Convert.ToDecimal(dr("U_HWSTE")),
1448 |                             .U_TAX_TYPE = dr("U_TAX_TYPE").ToString()
1449 |                         }
1450 |                         If Not IsDBNull(dr("U_BLDAT")) Then mdr.U_BLDAT = Convert.ToDateTime(dr("U_BLDAT"))
1451 |                         If Not IsDBNull(dr("U_VATDATE")) Then mdr.U_VATDATE = Convert.ToDateTime(dr("U_VATDATE"))
1452 |                         mdrLines.Add(mdr)
1453 |                     End While
1454 |                 End Using
1455 |             End Using
1456 |             CurrentMDRLines = mdrLines
1457 |             BindMDRGrid()
1458 | 
1459 |         End Using
1460 |     End Sub
1461 | #End Region
1462 | 
1463 | #Region "附件處理"
1464 |     Protected Sub btnUpload_Click(sender As Object, e As EventArgs)
1465 |         If fileUpload.HasFiles Then
1466 |             Try
1467 |                 Dim folder As String = Server.MapPath("~/Uploads/Expense/")
1468 |                 If Not Directory.Exists(folder) Then Directory.CreateDirectory(folder)
1469 | 
1470 |                 ' 若尚未存檔 (currentDocEntry=0)，是否需要先存草稿？
1471 |                 ' 策略：先存入 CurrentAttachments (ViewState/List)，待按下儲存/送出時寫入 jAttach
1472 |                 ' 但使用者可能希望上傳後就看得到檔案，且 jAttach 需要 jID (DocEntry)。
1473 |                 ' 方案：若 currentDocEntry=0，上傳時先只更新 ViewState，不寫入 DB。
1474 |                 ' 若 currentDocEntry > 0，直接寫入 DB 並更新 ViewState。
1475 |                 ' 但為了統一邏輯，這裡簡化為：上傳即存檔 (Draft)。若為新單，先執行 SaveDocument("P") 取得 DocEntry。
1476 |                 
1477 |                 If currentDocEntry = 0 Then
1478 |                     ' 自動儲存為草稿以取得 DocEntry
1479 |                     If Not SaveDocument("P", True) Then
1480 |                         ShowError("上傳附件前自動儲存草稿失敗，請檢查必填欄位。")
1481 |                         Return
1482 |                     End If
1483 |                 End If
1484 | 
1485 |                 Dim list = CurrentAttachments
1486 |                 Dim successCount As Integer = 0
1487 | 
1488 |                 For Each uploadedFile As HttpPostedFile In fileUpload.PostedFiles
1489 |                     Dim originName As String = Path.GetFileName(uploadedFile.FileName)
1490 |                     Dim savedName As String = DateTime.Now.ToString("yyyyMMddHHmmss") & "_" & originName
1491 |                     Dim fullPath As String = Path.Combine(folder, savedName)
1492 |                     uploadedFile.SaveAs(fullPath)
1493 | 
1494 |                     ' 寫入資料庫 jAttach
1495 |                     Using conn As New SqlConnection(connStr)
1496 |                         conn.Open()
1497 |                         Dim sql As String = "INSERT INTO jAttach (jID, DocEntry, LineNum, FilePath, FileName, Uploader, UploadTime) " &
1498 |                                           "VALUES (@jID, @DocEntry, -1, @FilePath, @FileName, @Uploader, @UploadTime); SELECT SCOPE_IDENTITY();"
1499 |                         Using cmd As New SqlCommand(sql, conn)
1500 |                             cmd.Parameters.AddWithValue("@jID", currentDocEntry) ' jOPCH.jID (DocEntry)
1501 |                             cmd.Parameters.AddWithValue("@DocEntry", currentDocEntry)
1502 |                             cmd.Parameters.AddWithValue("@FilePath", savedName) ' 存檔名作為路徑/識別
1503 |                             cmd.Parameters.AddWithValue("@FileName", originName)
1504 |                             cmd.Parameters.AddWithValue("@Uploader", currentUserId)
1505 |                             cmd.Parameters.AddWithValue("@UploadTime", DateTime.Now.ToString("HH:mm:ss"))
1506 |                             
1507 |                             Dim newId As Integer = Convert.ToInt32(cmd.ExecuteScalar())
1508 |                             
1509 |                             list.Add(New AttachmentItem With {
1510 |                                 .ID = newId,
1511 |                                 .FileName = savedName, ' 顯示與下載用存檔名，或可改顯示原檔名
1512 |                                 .FilePath = savedName,
1513 |                                 .UploadDate = DateTime.Today,
1514 |                                 .UploadTime = DateTime.Now.ToString("HH:mm:ss"),
1515 |                                 .Uploader = currentUserId
1516 |                             })
1517 |                             successCount += 1
1518 |                         End Using
1519 |                     End Using
1520 |                 Next
1521 |                 
1522 |                 CurrentAttachments = list
1523 |                 BindAttachmentGrid()
1524 |                 lblMessage.Text = $"成功上傳 {successCount} 個檔案"
1525 |                 lblMessage.ForeColor = System.Drawing.Color.Blue
1526 | 
1527 |             Catch ex As Exception
1528 |                 ShowError("上傳失敗: " & ex.Message)
1529 |             End Try
1530 |         Else
1531 |              ShowError("請選擇檔案")
1532 |         End If
1533 |     End Sub
1534 | 
1535 |     Protected Sub gvAttachments_RowCommand(sender As Object, e As GridViewCommandEventArgs)
1536 |         If e.CommandName = "DeleteFile" Then
1537 |             Dim index As Integer = Convert.ToInt32(e.CommandArgument)
1538 |             Dim list = CurrentAttachments
1539 |             
1540 |             If index < list.Count Then
1541 |                 Dim item = list(index)
1542 |                 
1543 |                 Try
1544 |                     ' 刪除 DB
1545 |                     Using conn As New SqlConnection(connStr)
1546 |                         conn.Open()
1547 |                         Dim sql As String = "DELETE FROM jAttach WHERE ID=@ID"
1548 |                         Using cmd As New SqlCommand(sql, conn)
1549 |                             cmd.Parameters.AddWithValue("@ID", item.ID)
1550 |                             cmd.ExecuteNonQuery()
1551 |                         End Using
1552 |                     End Using
1553 |                     
1554 |                     ' 刪除實體檔案 (可選，通常保留備份或標記刪除，這裡先不刪實體檔以防誤刪)
1555 |                     ' Dim path As String = Server.MapPath("~/Uploads/Expense/" & item.FilePath)
1556 |                     ' If File.Exists(path) Then File.Delete(path)
1557 | 
1558 |                     list.RemoveAt(index)
1559 |                     CurrentAttachments = list
1560 |                     BindAttachmentGrid()
1561 |                     lblMessage.Text = "附件已刪除"
1562 |                 Catch ex As Exception
1563 |                     ShowError("刪除附件失敗: " & ex.Message)
1564 |                 End Try
1565 |             End If
1566 |         End If
1567 |     End Sub
1568 | 
1569 |     Private Sub BindAttachmentGrid()
1570 |         gvAttachments.DataSource = CurrentAttachments
1571 |         gvAttachments.DataBind()
1572 |     End Sub
1573 | #End Region
1574 | 
1575 | #Region "審核"
1576 |     Protected Sub btnApprove_Click(sender As Object, e As EventArgs)
1577 |         If Not isApUser Then
1578 |             ShowError("無權限執行此操作")
1579 |             Return
1580 |         End If
1581 |         UpdateStatus("A")
1582 |     End Sub
1583 | 
1584 |     Protected Sub btnReject_Click(sender As Object, e As EventArgs)
1585 |         If Not isApUser Then
1586 |             ShowError("無權限執行此操作")
1587 |             Return
1588 |         End If
1589 |         UpdateStatus("R")
1590 |     End Sub
1591 | 
1592 |     Protected Sub btnUpdateComment_Click(sender As Object, e As EventArgs)
1593 |         If Not isApUser Then
1594 |             ShowError("無權限執行此操作")
1595 |             Return
1596 |         End If
1597 |         ' 僅更新意見，不改狀態
1598 |         Try
1599 |             Using conn As New SqlConnection(connStr)
1600 |                 conn.Open()
1601 |                 Dim sql As String = "UPDATE jOPCH SET ApprovalComments=@Comm, UpdateBy=@User, UpdateDate=GETDATE() WHERE DocEntry=@ID"
1602 |                 Using cmd As New SqlCommand(sql, conn)
1603 |                     cmd.Parameters.AddWithValue("@User", currentUserId)
1604 |                     cmd.Parameters.AddWithValue("@Comm", txtApprovalComments.Text)
1605 |                     cmd.Parameters.AddWithValue("@ID", currentDocEntry)
1606 |                     cmd.ExecuteNonQuery()
1607 |                 End Using
1608 |             End Using
1609 |             ShowError("意見已更新") ' 用紅色顯示有點怪，但先這樣
1610 |             lblMessage.ForeColor = System.Drawing.Color.Blue
1611 |         Catch ex As Exception
1612 |             ShowError("更新失敗: " & ex.Message)
1613 |         End Try
1614 |     End Sub
1615 | 
1616 |     Private Sub UpdateStatus(status As String)
1617 |         Try
1618 |             Using conn As New SqlConnection(connStr)
1619 |                 conn.Open()
1620 |                 Dim sql As String = "UPDATE jOPCH SET ApprovalStatus=@Status, ApprovedBy=@User, ApprovalDate=GETDATE(), ApprovalComments=@Comm WHERE DocEntry=@ID"
1621 |                 Using cmd As New SqlCommand(sql, conn)
1622 |                     cmd.Parameters.AddWithValue("@Status", status)
1623 |                     cmd.Parameters.AddWithValue("@User", currentUserId)
1624 |                     cmd.Parameters.AddWithValue("@Comm", txtApprovalComments.Text)
1625 |                     cmd.Parameters.AddWithValue("@ID", currentDocEntry)
1626 |                     cmd.ExecuteNonQuery()
1627 |                 End Using
1628 |             End Using
1629 | 
1630 |             ' 如果是放行 (A)，這裡應呼叫 SAP B1 API 建立 AP Invoice
1631 |             If status = "A" Then
1632 |                 ' TODO: Call SAP B1 API & MDR Integration
1633 |                 ' CreateAPInvoiceInSAP(currentDocEntry)
1634 |             End If
1635 | 
1636 |             Response.Redirect("ExpenseClaimForm.aspx?DocEntry=" & currentDocEntry)
1637 |         Catch ex As Exception
1638 |             ShowError("更新狀態失敗: " & ex.Message)
1639 |         End Try
1640 |     End Sub
1641 | #End Region
1642 | #Region "輔助函式"
1643 |     Private Sub ShowError(msg As String)
1644 |         lblMessage.Text = msg
1645 |         lblMessage.ForeColor = System.Drawing.Color.Red
1646 |     End Sub
1647 | 
1648 |     Private Function GetVatRate(vatCode As String) As Decimal
1649 |         If vatCode = "1" Then Return 5
1650 |         Return 0
1651 |     End Function
1652 | 
1653 |     Private Function GetStatusText(status As String) As String
1654 |         Select Case status
1655 |             Case "P" : Return "草稿"
1656 |             Case "W" : Return "待審核"
1657 |             Case "A" : Return "已核准"
1658 |             Case "R" : Return "駁回"
1659 |             Case Else : Return status
1660 |         End Select
1661 |     End Function
1662 | #End Region
1663 | 
1664 | End Class
</content>
</file>
<file><path>shopfloor/Claude_TMP/etc/roo-code-encoding-fix.md</path>
<content lines="1-334">
  1 | # Roo Code 執行指令：解決 apply_diff 編碼失敗問題
  2 | 
  3 | ## 問題描述
  4 | 
  5 | `apply_diff` 失敗且相似度極高 (93%-99%)，原因是檔案編碼 (BOM) 與換行符號 (CRLF/LF) 的隱形差異。
  6 | 
  7 | ## 解決方案
  8 | 
  9 | 統一專案檔案格式為：**UTF-8 with BOM + LF**
 10 | 
 11 | ---
 12 | 
 13 | ## 步驟 0：備份現有檔案（必須先執行）
 14 | 
 15 | ### 0.1 建立備份資料夾並複製所有目標檔案
 16 | 
 17 | ```powershell
 18 | # 設定變數
 19 | $projectDir = "."  # 改成你的專案路徑
 20 | $backupDir = ".\_encoding_backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
 21 | $extensions = @("*.vb", "*.aspx", "*.ascx", "*.asax", "*.master", "*.config")
 22 | 
 23 | # 建立備份資料夾
 24 | New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
 25 | Write-Host "備份資料夾: $backupDir" -ForegroundColor Cyan
 26 | 
 27 | # 備份檔案（保留目錄結構）
 28 | $count = 0
 29 | foreach ($ext in $extensions) {
 30 |     Get-ChildItem -Path $projectDir -Filter $ext -Recurse | ForEach-Object {
 31 |         $relativePath = $_.FullName.Substring((Get-Item $projectDir).FullName.Length + 1)
 32 |         $destPath = Join-Path $backupDir $relativePath
 33 |         $destDir = Split-Path $destPath -Parent
 34 |         
 35 |         if (!(Test-Path $destDir)) {
 36 |             New-Item -ItemType Directory -Path $destDir -Force | Out-Null
 37 |         }
 38 |         
 39 |         Copy-Item $_.FullName -Destination $destPath
 40 |         $count++
 41 |     }
 42 | }
 43 | 
 44 | Write-Host "已備份 $count 個檔案到 $backupDir" -ForegroundColor Green
 45 | 
 46 | # 產生備份清單
 47 | Get-ChildItem -Path $backupDir -Recurse -File | 
 48 |     Select-Object FullName, Length, LastWriteTime |
 49 |     Export-Csv -Path "$backupDir\_backup_manifest.csv" -NoTypeInformation -Encoding UTF8
 50 | 
 51 | Write-Host "備份清單已儲存: $backupDir\_backup_manifest.csv" -ForegroundColor Green
 52 | ```
 53 | 
 54 | ### 0.2 驗證備份完整性
 55 | 
 56 | ```powershell
 57 | # 確認備份檔案數量
 58 | $backupDir = ".\_encoding_backup_*" | Get-Item | Sort-Object Name -Descending | Select-Object -First 1
 59 | $backupCount = (Get-ChildItem -Path $backupDir -Recurse -File | Where-Object { $_.Name -ne "_backup_manifest.csv" }).Count
 60 | Write-Host "備份檔案數量: $backupCount" -ForegroundColor Cyan
 61 | 
 62 | # 確認備份資料夾存在且不為空才繼續
 63 | if ($backupCount -gt 0) {
 64 |     Write-Host "✓ 備份完成，可以繼續執行標準化" -ForegroundColor Green
 65 | } else {
 66 |     Write-Host "✗ 備份失敗，請勿繼續執行！" -ForegroundColor Red
 67 |     exit 1
 68 | }
 69 | ```
 70 | 
 71 | ---
 72 | 
 73 | ## 步驟 1：建立 .editorconfig（專案根目錄）
 74 | 
 75 | ```ini
 76 | # EditorConfig - 統一編碼格式
 77 | root = true
 78 | 
 79 | [*]
 80 | indent_style = space
 81 | indent_size = 4
 82 | end_of_line = lf
 83 | trim_trailing_whitespace = true
 84 | insert_final_newline = true
 85 | 
 86 | [*.{vb,aspx,ascx,asax,master,config}]
 87 | charset = utf-8-bom
 88 | end_of_line = lf
 89 | ```
 90 | 
 91 | ---
 92 | 
 93 | ## 步驟 2：建立 .gitattributes（專案根目錄）
 94 | 
 95 | ```
 96 | # 防止 Git 自動轉換換行符
 97 | * text=auto eol=lf
 98 | 
 99 | # ASP.NET 專案檔案
100 | *.vb text eol=lf
101 | *.aspx text eol=lf
102 | *.ascx text eol=lf
103 | *.asax text eol=lf
104 | *.master text eol=lf
105 | *.config text eol=lf
106 | *.cs text eol=lf
107 | ```
108 | 
109 | ---
110 | 
111 | ## 步驟 3：一次性標準化現有檔案
112 | 
113 | 執行以下 PowerShell 腳本，將所有目標檔案轉換為「UTF-8 with BOM + LF」：
114 | 
115 | ```powershell
116 | # ============================================
117 | # 標準化腳本：UTF-8 BOM + LF
118 | # ============================================
119 | 
120 | $extensions = @("*.vb", "*.aspx", "*.ascx", "*.asax", "*.master", "*.config")
121 | $targetDir = "."  # 改成你的專案路徑
122 | 
123 | $successCount = 0
124 | $errorCount = 0
125 | $errorFiles = @()
126 | 
127 | foreach ($ext in $extensions) {
128 |     Get-ChildItem -Path $targetDir -Filter $ext -Recurse | ForEach-Object {
129 |         $file = $_.FullName
130 |         
131 |         try {
132 |             # 讀取內容（自動處理各種編碼）
133 |             $content = [System.IO.File]::ReadAllText($file)
134 |             
135 |             # 統一換行符為 LF
136 |             $content = $content -replace "`r`n", "`n"
137 |             $content = $content -replace "`r", "`n"
138 |             
139 |             # 寫入 UTF-8 with BOM
140 |             $utf8Bom = New-Object System.Text.UTF8Encoding($true)
141 |             [System.IO.File]::WriteAllText($file, $content, $utf8Bom)
142 |             
143 |             Write-Host "✓ $file" -ForegroundColor Green
144 |             $successCount++
145 |         }
146 |         catch {
147 |             Write-Host "✗ $file - $($_.Exception.Message)" -ForegroundColor Red
148 |             $errorCount++
149 |             $errorFiles += $file
150 |         }
151 |     }
152 | }
153 | 
154 | Write-Host "`n========================================" -ForegroundColor Cyan
155 | Write-Host "轉換完成！" -ForegroundColor Cyan
156 | Write-Host "成功: $successCount 個檔案" -ForegroundColor Green
157 | Write-Host "失敗: $errorCount 個檔案" -ForegroundColor $(if ($errorCount -gt 0) { "Red" } else { "Green" })
158 | 
159 | if ($errorFiles.Count -gt 0) {
160 |     Write-Host "`n失敗的檔案:" -ForegroundColor Red
161 |     $errorFiles | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
162 | }
163 | ```
164 | 
165 | ---
166 | 
167 | ## 步驟 4：驗證轉換結果
168 | 
169 | ```powershell
170 | # ============================================
171 | # 驗證腳本：檢查所有檔案格式
172 | # ============================================
173 | 
174 | $extensions = @("*.vb", "*.aspx", "*.ascx", "*.asax", "*.master", "*.config")
175 | $targetDir = "."
176 | 
177 | $correctFiles = 0
178 | $incorrectFiles = @()
179 | 
180 | foreach ($ext in $extensions) {
181 |     Get-ChildItem -Path $targetDir -Filter $ext -Recurse | ForEach-Object {
182 |         $file = $_.FullName
183 |         $bytes = [System.IO.File]::ReadAllBytes($file)
184 |         
185 |         # 檢查 BOM
186 |         $hasBom = ($bytes.Length -ge 3) -and 
187 |                   ($bytes[0] -eq 0xEF) -and 
188 |                   ($bytes[1] -eq 0xBB) -and 
189 |                   ($bytes[2] -eq 0xBF)
190 |         
191 |         # 檢查換行符（跳過 BOM 後的內容）
192 |         $content = [System.IO.File]::ReadAllText($file)
193 |         $hasCRLF = $content.Contains("`r`n")
194 |         
195 |         if ($hasBom -and !$hasCRLF) {
196 |             $correctFiles++
197 |         } else {
198 |             $issues = @()
199 |             if (!$hasBom) { $issues += "缺少 BOM" }
200 |             if ($hasCRLF) { $issues += "含有 CRLF" }
201 |             $incorrectFiles += [PSCustomObject]@{
202 |                 File = $file
203 |                 Issues = $issues -join ", "
204 |             }
205 |         }
206 |     }
207 | }
208 | 
209 | Write-Host "`n========================================" -ForegroundColor Cyan
210 | Write-Host "驗證結果" -ForegroundColor Cyan
211 | Write-Host "========================================" -ForegroundColor Cyan
212 | Write-Host "正確格式: $correctFiles 個檔案" -ForegroundColor Green
213 | 
214 | if ($incorrectFiles.Count -gt 0) {
215 |     Write-Host "格式錯誤: $($incorrectFiles.Count) 個檔案" -ForegroundColor Red
216 |     Write-Host "`n問題檔案:" -ForegroundColor Yellow
217 |     $incorrectFiles | ForEach-Object {
218 |         Write-Host "  - $($_.File)" -ForegroundColor Red
219 |         Write-Host "    問題: $($_.Issues)" -ForegroundColor Yellow
220 |     }
221 | } else {
222 |     Write-Host "`n✓ 所有檔案格式正確！" -ForegroundColor Green
223 | }
224 | ```
225 | 
226 | ---
227 | 
228 | ## 步驟 5：提交變更到 Git
229 | 
230 | ```bash
231 | git add .editorconfig .gitattributes
232 | git add -A
233 | git commit -m "chore: 標準化檔案編碼為 UTF-8 BOM + LF"
234 | ```
235 | 
236 | ---
237 | 
238 | ## Rollback：還原到備份狀態
239 | 
240 | 如果轉換後出現問題（亂碼、編譯錯誤等），執行以下腳本還原：
241 | 
242 | ```powershell
243 | # ============================================
244 | # Rollback 腳本：從備份還原
245 | # ============================================
246 | 
247 | # 找到最新的備份資料夾
248 | $backupDir = Get-ChildItem -Path "." -Directory -Filter "_encoding_backup_*" | 
249 |              Sort-Object Name -Descending | 
250 |              Select-Object -First 1
251 | 
252 | if ($null -eq $backupDir) {
253 |     Write-Host "✗ 找不到備份資料夾！" -ForegroundColor Red
254 |     exit 1
255 | }
256 | 
257 | Write-Host "找到備份: $($backupDir.FullName)" -ForegroundColor Cyan
258 | Write-Host "準備還原..." -ForegroundColor Yellow
259 | 
260 | # 確認還原
261 | $confirm = Read-Host "確定要還原嗎？這會覆蓋目前的檔案 (y/N)"
262 | if ($confirm -ne "y" -and $confirm -ne "Y") {
263 |     Write-Host "已取消還原" -ForegroundColor Yellow
264 |     exit 0
265 | }
266 | 
267 | # 執行還原
268 | $restoreCount = 0
269 | $errorCount = 0
270 | 
271 | Get-ChildItem -Path $backupDir.FullName -Recurse -File | 
272 |     Where-Object { $_.Name -ne "_backup_manifest.csv" } | 
273 |     ForEach-Object {
274 |         $relativePath = $_.FullName.Substring($backupDir.FullName.Length + 1)
275 |         $destPath = Join-Path "." $relativePath
276 |         
277 |         try {
278 |             # 確保目標目錄存在
279 |             $destDir = Split-Path $destPath -Parent
280 |             if (!(Test-Path $destDir)) {
281 |                 New-Item -ItemType Directory -Path $destDir -Force | Out-Null
282 |             }
283 |             
284 |             Copy-Item $_.FullName -Destination $destPath -Force
285 |             Write-Host "✓ 還原: $relativePath" -ForegroundColor Green
286 |             $restoreCount++
287 |         }
288 |         catch {
289 |             Write-Host "✗ 還原失敗: $relativePath - $($_.Exception.Message)" -ForegroundColor Red
290 |             $errorCount++
291 |         }
292 |     }
293 | 
294 | Write-Host "`n========================================" -ForegroundColor Cyan
295 | Write-Host "還原完成！" -ForegroundColor Cyan
296 | Write-Host "成功: $restoreCount 個檔案" -ForegroundColor Green
297 | Write-Host "失敗: $errorCount 個檔案" -ForegroundColor $(if ($errorCount -gt 0) { "Red" } else { "Green" })
298 | 
299 | # 提醒刪除設定檔
300 | Write-Host "`n如果要完全還原，請手動刪除以下檔案:" -ForegroundColor Yellow
301 | Write-Host "  - .editorconfig" -ForegroundColor Yellow
302 | Write-Host "  - .gitattributes" -ForegroundColor Yellow
303 | ```
304 | 
305 | ---
306 | 
307 | ## 後續 apply_diff 規則
308 | 
309 | 完成上述設定後，在 Roo Code 的 Custom Instructions 或 `.roo/rules` 中加入：
310 | 
311 | ```markdown
312 | ## 專案編碼規範
313 | 
314 | 此專案的 .vb / .aspx / .ascx 等檔案格式為：
315 | - 編碼：UTF-8 with BOM
316 | - 換行符：LF
317 | 
318 | 執行 apply_diff 時：
319 | 1. 確保輸出內容的換行符使用 LF（\n），不要使用 CRLF（\r\n）
320 | 2. 不要移除檔案開頭的 BOM (EF BB BF)
321 | ```
322 | 
323 | ---
324 | 
325 | ## 執行順序檢查清單
326 | 
327 | - [ ] 步驟 0：執行備份並確認備份完整
328 | - [ ] 步驟 1：建立 .editorconfig
329 | - [ ] 步驟 2：建立 .gitattributes  
330 | - [ ] 步驟 3：執行標準化腳本
331 | - [ ] 步驟 4：執行驗證腳本確認全部正確
332 | - [ ] 測試：部署到測試環境確認沒有亂碼
333 | - [ ] 步驟 5：提交到 Git
334 | - [ ] （如有問題）執行 Rollback 腳本還原
</content>
</file>
</files>
<environment_details>
# VSCode Visible Files
MgmSP\ExpenseClaimForm.aspx.vb

# VSCode Open Tabs
shopfloor/Claude_TMP/規格差異與實作方向確認_20251125.md,shopfloor/Claude_TMP/ExpenseClaimForm_Spec_Integration_20251126.md,DevExp/global/ui-ux.toml,agent-os/standards/global/localization.md,agent-os/config.yml,mcp-sqlserver/config.json,agent-os/standards/global/environment-health-check.md,agent-os/SESSION_INIT.md,DevExp/global/environment.toml,.git/COMMIT_EDITMSG,MgmSP/ExpenseClaimForm_DataSpec_20251203.md,MgmSP/ExpenseClaimForm.aspx,DevExp/projects/SapB1Solution/specific.toml,.editorconfig,.gitattributes,MgmSP/ExpenseClaimForm.aspx.vb,MgmSP/ExpenseClaimForm.aspx.designer.vb

# Current Time
Current time in ISO 8601 UTC format: 2025-12-04T04:02:17.338Z
User time zone: Asia/Taipei, UTC+8:00

# Current Cost
$0.00

# Current Mode
<slug>code</slug>
<name>💻 Code</name>
<model>claude-opus-4-5-20251101</model>
You have not created a todo list yet. Create one with `update_todo_list` if your task is complicated or involves multiple steps.
</environment_details>

---

**Assistant:**

我需要讀取更多檔案來完成分析。讓我先建立一個待辦清單，然後繼續讀取相關檔案。

---

**User:**

[ERROR] You did not use a tool in your previous response! Please retry with a tool use.

# Reminder: Instructions for Tool Use

Tool uses are formatted using XML-style tags. The tool name itself becomes the XML tag name. Each parameter is enclosed within its own set of tags. Here's the structure:

<actual_tool_name>
<parameter1_name>value1</parameter1_name>
<parameter2_name>value2</parameter2_name>
...
</actual_tool_name>

For example, to use the attempt_completion tool:

<attempt_completion>
<result>
I have completed the task...
</result>
</attempt_completion>

Always use the actual tool name as the XML tag name for proper parsing and execution.

# Next Steps

If you have completed the user's task, use the attempt_completion tool.
If you require additional information from the user, use the ask_followup_question tool.
Otherwise, if you have not completed the task and do not need additional information, then proceed with the next step of the task.
(This is an automated message, so do not respond to it conversationally.)
<environment_details>
# VSCode Visible Files
MgmSP\ExpenseClaimForm.aspx.vb

# VSCode Open Tabs
shopfloor/Claude_TMP/規格差異與實作方向確認_20251125.md,shopfloor/Claude_TMP/ExpenseClaimForm_Spec_Integration_20251126.md,DevExp/global/ui-ux.toml,agent-os/standards/global/localization.md,agent-os/config.yml,mcp-sqlserver/config.json,agent-os/standards/global/environment-health-check.md,agent-os/SESSION_INIT.md,DevExp/global/environment.toml,.git/COMMIT_EDITMSG,MgmSP/ExpenseClaimForm_DataSpec_20251203.md,MgmSP/ExpenseClaimForm.aspx,DevExp/projects/SapB1Solution/specific.toml,.editorconfig,.gitattributes,MgmSP/ExpenseClaimForm.aspx.vb,MgmSP/ExpenseClaimForm.aspx.designer.vb

# Current Time
Current time in ISO 8601 UTC format: 2025-12-04T04:02:57.307Z
User time zone: Asia/Taipei, UTC+8:00

# Current Cost
$0.00

# Current Mode
<slug>code</slug>
<name>💻 Code</name>
<model>claude-opus-4-5-20251101</model>
You have not created a todo list yet. Create one with `update_todo_list` if your task is complicated or involves multiple steps.
</environment_details>

