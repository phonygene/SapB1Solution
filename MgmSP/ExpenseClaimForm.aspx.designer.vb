'------------------------------------------------------------------------------
' <自動產生的>
'     這段程式碼是由工具產生的。
'
'     變更這個檔案可能會導致不正確的行為，而且如果已重新產生
'     程式碼，則會遺失變更。
' </自動產生的>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On


Partial Public Class ExpenseClaimForm

    '''<summary>
    '''form1 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents form1 As Global.System.Web.UI.HtmlControls.HtmlForm

    '''<summary>
    '''ScriptManager1 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents ScriptManager1 As Global.System.Web.UI.ScriptManager

    '''<summary>
    '''hfActiveTab 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents hfActiveTab As Global.System.Web.UI.WebControls.HiddenField

    '''<summary>
    '''UpdatePanel1 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents UpdatePanel1 As Global.System.Web.UI.UpdatePanel

    '''<summary>
    '''lblDocNum 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents lblDocNum As Global.System.Web.UI.WebControls.Label

    '''<summary>
    '''lblDocStatus 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents lblDocStatus As Global.System.Web.UI.WebControls.Label

    '''<summary>
    '''txtCardCode 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents txtCardCode As Global.System.Web.UI.WebControls.TextBox

    '''<summary>
    '''btnSearchCardCode 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents btnSearchCardCode As Global.System.Web.UI.WebControls.Button

    '''<summary>
    '''lblVendorInfo 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents lblVendorInfo As Global.System.Web.UI.WebControls.Label

    '''<summary>
    '''lblErrCardCode 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents lblErrCardCode As Global.System.Web.UI.WebControls.Label

    '''<summary>
    '''txtCardName 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents txtCardName As Global.System.Web.UI.WebControls.TextBox

    '''<summary>
    '''btnSearchCardName 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents btnSearchCardName As Global.System.Web.UI.WebControls.Button

    '''<summary>
    '''lblErrCardName 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents lblErrCardName As Global.System.Web.UI.WebControls.Label

    '''<summary>
    '''txtNumAtCard 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents txtNumAtCard As Global.System.Web.UI.WebControls.TextBox

    '''<summary>
    '''ddlDocCurrency 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents ddlDocCurrency As Global.System.Web.UI.WebControls.DropDownList

    '''<summary>
    '''txtDocRate 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents txtDocRate As Global.System.Web.UI.WebControls.TextBox

    '''<summary>
    '''btnRefreshRate 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents btnRefreshRate As Global.System.Web.UI.WebControls.Button

    '''<summary>
    '''ddlDeliveryAddr 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents ddlDeliveryAddr As Global.System.Web.UI.WebControls.DropDownList

    '''<summary>
    '''txtAddress 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents txtAddress As Global.System.Web.UI.WebControls.TextBox

    '''<summary>
    '''ddlGroupNum 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents ddlGroupNum As Global.System.Web.UI.WebControls.DropDownList

    '''<summary>
    '''txtJID 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents txtJID As Global.System.Web.UI.WebControls.TextBox

    '''<summary>
    '''txtB1DocEntry 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents txtB1DocEntry As Global.System.Web.UI.WebControls.TextBox

    '''<summary>
    '''txtUPID 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents txtUPID As Global.System.Web.UI.WebControls.TextBox

    '''<summary>
    '''txtStatusDisplay 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents txtStatusDisplay As Global.System.Web.UI.WebControls.TextBox

    '''<summary>
    '''txtTaxDate 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents txtTaxDate As Global.System.Web.UI.WebControls.TextBox

    '''<summary>
    '''txtDocDueDate 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents txtDocDueDate As Global.System.Web.UI.WebControls.TextBox

    '''<summary>
    '''lblErrDocDueDate 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents lblErrDocDueDate As Global.System.Web.UI.WebControls.Label

    '''<summary>
    '''txtDocDate 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents txtDocDate As Global.System.Web.UI.WebControls.TextBox

    '''<summary>
    '''lblErrDocDate 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents lblErrDocDate As Global.System.Web.UI.WebControls.Label

    '''<summary>
    '''txtApprovalStatus 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents txtApprovalStatus As Global.System.Web.UI.WebControls.TextBox

    '''<summary>
    '''txtApprovedBy 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents txtApprovedBy As Global.System.Web.UI.WebControls.TextBox

    '''<summary>
    '''btnTabExpense 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents btnTabExpense As Global.System.Web.UI.HtmlControls.HtmlButton

    '''<summary>
    '''btnTabMDR 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents btnTabMDR As Global.System.Web.UI.HtmlControls.HtmlButton

    '''<summary>
    '''divContentExpense 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents divContentExpense As Global.System.Web.UI.HtmlControls.HtmlGenericControl

    '''<summary>
    '''btnAddLine 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents btnAddLine As Global.System.Web.UI.WebControls.Button

    '''<summary>
    '''btnDeleteLine 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents btnDeleteLine As Global.System.Web.UI.WebControls.Button

    '''<summary>
    '''fileUpload 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents fileUpload As Global.System.Web.UI.WebControls.FileUpload

    '''<summary>
    '''btnUpload 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents btnUpload As Global.System.Web.UI.WebControls.Button

    '''<summary>
    '''gvAttachments 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents gvAttachments As Global.System.Web.UI.WebControls.GridView

    '''<summary>
    '''lblAttachment 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents lblAttachment As Global.System.Web.UI.WebControls.Label

    '''<summary>
    '''gvExpenseDetail 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents gvExpenseDetail As Global.System.Web.UI.WebControls.GridView

    '''<summary>
    '''divContentMDR 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents divContentMDR As Global.System.Web.UI.HtmlControls.HtmlGenericControl

    '''<summary>
    '''btnAddMDRRow 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents btnAddMDRRow As Global.System.Web.UI.WebControls.Button

    '''<summary>
    '''btnDeleteMDRRow 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents btnDeleteMDRRow As Global.System.Web.UI.WebControls.Button

    '''<summary>
    '''gvMDRDetail 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents gvMDRDetail As Global.System.Web.UI.WebControls.GridView

    '''<summary>
    '''ddlPurchaser 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents ddlPurchaser As Global.System.Web.UI.WebControls.DropDownList

    '''<summary>
    '''txtOwner 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents txtOwner As Global.System.Web.UI.WebControls.TextBox

    '''<summary>
    '''txtRemarks 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents txtRemarks As Global.System.Web.UI.WebControls.TextBox

    '''<summary>
    '''lblDocTotalWithTax 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents lblDocTotalWithTax As Global.System.Web.UI.WebControls.Label

    '''<summary>
    '''lblDocTotal 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents lblDocTotal As Global.System.Web.UI.WebControls.Label

    '''<summary>
    '''lblVatSum 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents lblVatSum As Global.System.Web.UI.WebControls.Label

    '''<summary>
    '''btnSave 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents btnSave As Global.System.Web.UI.WebControls.Button

    '''<summary>
    '''btnSubmit 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents btnSubmit As Global.System.Web.UI.WebControls.Button

    '''<summary>
    '''btnDelete 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents btnDelete As Global.System.Web.UI.WebControls.Button

    '''<summary>
    '''btnCancel 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents btnCancel As Global.System.Web.UI.WebControls.Button

    '''<summary>
    '''lblMessage 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents lblMessage As Global.System.Web.UI.WebControls.Label

    '''<summary>
    '''pnlApproval 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents pnlApproval As Global.System.Web.UI.WebControls.Panel

    '''<summary>
    '''txtApprovalComments 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents txtApprovalComments As Global.System.Web.UI.WebControls.TextBox

    '''<summary>
    '''btnApprove 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents btnApprove As Global.System.Web.UI.WebControls.Button

    '''<summary>
    '''btnUpdateComment 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents btnUpdateComment As Global.System.Web.UI.WebControls.Button

    '''<summary>
    '''btnReject 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents btnReject As Global.System.Web.UI.WebControls.Button

    '''<summary>
    '''btnDummy 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents btnDummy As Global.System.Web.UI.WebControls.Button

    '''<summary>
    '''mpeVendor 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents mpeVendor As Global.AjaxControlToolkit.ModalPopupExtender

    '''<summary>
    '''pnlVendorSearch 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents pnlVendorSearch As Global.System.Web.UI.WebControls.Panel

    '''<summary>
    '''btnCloseVendor 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents btnCloseVendor As Global.System.Web.UI.WebControls.LinkButton

    '''<summary>
    '''txtVendorSearchKeyword 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents txtVendorSearchKeyword As Global.System.Web.UI.WebControls.TextBox

    '''<summary>
    '''btnDoSearchVendor 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents btnDoSearchVendor As Global.System.Web.UI.WebControls.Button

    '''<summary>
    '''hfSearchSource 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents hfSearchSource As Global.System.Web.UI.WebControls.HiddenField

    '''<summary>
    '''rblSearchMode 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents rblSearchMode As Global.System.Web.UI.WebControls.RadioButtonList

    '''<summary>
    '''gvVendorSearch 控制項。
    '''</summary>
    '''<remarks>
    '''自動產生的欄位。
    '''若要修改，請將欄位宣告從設計工具檔案移到程式碼後置檔案。
    '''</remarks>
    Protected WithEvents gvVendorSearch As Global.System.Web.UI.WebControls.GridView
End Class
