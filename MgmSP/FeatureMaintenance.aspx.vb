Imports System.Data.SqlClient
Imports System.Web.Configuration

''' <summary>
''' 功能維護頁面
''' 當特定功能維護中時，使用者會被導向此頁面
''' </summary>
Partial Public Class FeatureMaintenance
    Inherits System.Web.UI.Page

    Private ReadOnly connStr As String = WebConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            LoadFeatureInfo()
        End If
    End Sub

    ''' <summary>
    ''' 載入功能資訊
    ''' </summary>
    Private Sub LoadFeatureInfo()
        Dim featureName As String = Request.QueryString("feature")
        
        ' 顯示功能名稱
        Select Case featureName
            Case "ExpenseClaim"
                litFeatureName.Text = "費用申請單"
            Case "PurchaseRequest"
                litFeatureName.Text = "請購單"
            Case Else
                litFeatureName.Text = "此功能"
        End Select
        
        ' 載入維護訊息
        litMaintenanceNote.Text = Server.HtmlEncode(MaintenanceHelper.GetMaintenanceMessage()).Replace(vbCrLf, "<br/>").Replace(vbLf, "<br/>")
    End Sub

End Class
