# Claude Code 專案配置

## 專案概述
這是一個 ASP.NET Web Forms 專案 (VB.NET)，使用 Visual Studio 開發。

## 重要規則：ASP.NET Web Forms 控制項宣告

### **designer.vb 檔案規則 (非常重要!)**

當在 `.aspx` 檔案中新增控制項時，**必須同時更新對應的 `.aspx.designer.vb` 檔案**，否則會出現編譯錯誤：
- `'控制項名稱' 未宣告。由於其保護層級，可能無法對其進行存取。`

### 新增控制項的步驟：

1. **在 .aspx 檔案新增控制項**
   ```html
   <asp:Button ID="btnMyButton" runat="server" Text="按鈕" />
   <asp:DropDownList ID="ddlMyList" runat="server" />
   <ajaxToolkit:ModalPopupExtender ID="mpeMyModal" runat="server" ... />
   ```

2. **在 .aspx.designer.vb 檔案新增對應宣告**
   ```vb
   '''<summary>
   '''btnMyButton 控制項。
   '''</summary>
   Protected WithEvents btnMyButton As Global.System.Web.UI.WebControls.Button

   '''<summary>
   '''ddlMyList 控制項。
   '''</summary>
   Protected WithEvents ddlMyList As Global.System.Web.UI.WebControls.DropDownList

   '''<summary>
   '''mpeMyModal 控制項。
   '''</summary>
   Protected WithEvents mpeMyModal As Global.AjaxControlToolkit.ModalPopupExtender
   ```

### 常用控制項型別對照表：

| ASPX 控制項 | Designer.vb 型別 |
|------------|-----------------|
| `<asp:Button>` | `Global.System.Web.UI.WebControls.Button` |
| `<asp:TextBox>` | `Global.System.Web.UI.WebControls.TextBox` |
| `<asp:DropDownList>` | `Global.System.Web.UI.WebControls.DropDownList` |
| `<asp:Label>` | `Global.System.Web.UI.WebControls.Label` |
| `<asp:Literal>` | `Global.System.Web.UI.WebControls.Literal` |
| `<asp:Panel>` | `Global.System.Web.UI.WebControls.Panel` |
| `<asp:GridView>` | `Global.System.Web.UI.WebControls.GridView` |
| `<asp:HiddenField>` | `Global.System.Web.UI.WebControls.HiddenField` |
| `<asp:LinkButton>` | `Global.System.Web.UI.WebControls.LinkButton` |
| `<asp:CheckBox>` | `Global.System.Web.UI.WebControls.CheckBox` |
| `<asp:RadioButtonList>` | `Global.System.Web.UI.WebControls.RadioButtonList` |
| `<asp:FileUpload>` | `Global.System.Web.UI.WebControls.FileUpload` |
| `<asp:BulletedList>` | `Global.System.Web.UI.WebControls.BulletedList` |
| `<ajaxToolkit:ModalPopupExtender>` | `Global.AjaxControlToolkit.ModalPopupExtender` |
| `<div runat="server">` | `Global.System.Web.UI.HtmlControls.HtmlGenericControl` |
| `<button runat="server">` | `Global.System.Web.UI.HtmlControls.HtmlButton` |
| `<form runat="server">` | `Global.System.Web.UI.HtmlControls.HtmlForm` |

## Namespace 規則

此專案的 Root Namespace 是 `MgmSP`（定義在 .vbproj 中）。

### .vb 程式碼檔案
**不需要**在程式碼中明確宣告 Namespace，VB.NET 編譯器會自動將類別放在 `MgmSP` 命名空間下。

```vb
' 正確 - 不需要寫 Namespace
Imports System.Data.SqlClient

Public Class MyClass
    ' 編譯後會自動成為 MgmSP.MyClass
End Class
```

### .aspx 頁面的 Inherits 屬性
**必須**包含 Root Namespace 前綴：

```html
<%@ Page ... Inherits="MgmSP.MyClass" %>
```

**錯誤範例：**
```html
<%@ Page ... Inherits="MyClass" %>  <!-- 錯誤！會找不到類別 -->
```

## 檔案編碼規則 (非常重要!)

此專案所有檔案必須使用 **UTF-8 with BOM** 編碼，否則中文字會變成亂碼。

當建立新檔案時，必須在檔案開頭加入 BOM (Byte Order Mark)：
- BOM 的十六進位值: `EF BB BF`
- 在 Bash 中可用: `printf '\xEF\xBB\xBF' > newfile.aspx`

**檢查檔案編碼的方式：**
```bash
file "檔案路徑"
# 正確: UTF-8 (with BOM) text
# 錯誤: UTF-8 text (沒有 BOM)
```

**修復缺少 BOM 的檔案：**
```bash
printf '\xEF\xBB\xBF' > file.tmp && cat original_file >> file.tmp && mv file.tmp original_file
```

## 新增 .aspx 頁面的完整步驟

1. 建立 `PageName.aspx` - HTML/ASPX 標記 **(必須 UTF-8 with BOM)**
2. 建立 `PageName.aspx.vb` - 程式碼後置檔案 **(必須 UTF-8 with BOM)**
3. 建立 `PageName.aspx.designer.vb` - 控制項宣告檔案 **(必須 UTF-8 with BOM)**
4. 確認 `<%@ Page ... Inherits="MgmSP.ClassName" %>` 中 **必須**加上 `MgmSP.` 前綴
5. 重新建置專案

## 資料庫連線

- 本地資料庫連線字串名稱: `jtdbConnectionString`
- SAP 資料庫連線字串名稱: `SapSQLConnection`

## 常用輔助類別

- `CommUtil` - 通用工具類別，包含資料庫查詢方法
- `MaintenanceHelper` - 系統維護檢查模組
