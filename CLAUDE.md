# Claude Code 專案配置

## 專案概述
這是一個 ASP.NET Web Forms 專案 (VB.NET)，使用 Visual Studio 開發。
**此為財務相關系統，資料正確性極為重要，錯誤可能導致帳務與法律問題。**

## 開發核心原則 (必須遵守!)

開發時必須遵循以下原則，確保用戶看到的資料與實際儲存的資料完全一致：

| 原則 | 說明 | 實踐方式 |
|------|------|----------|
| **POLA** | 最小驚訝原則 - 系統行為應符合用戶預期，不應有意外結果 | 不在背景偷偷修改用戶輸入的值 |
| **WYSIWYG** | 所見即所得 - 畫面顯示什麼，就儲存什麼 | 儲存前不重新計算已顯示的值 |
| **Data Consistency** | 資料一致性 - 輸入與儲存的資料應一致 | 避免在 Save 時修改 Model 的值 |
| **Form State Integrity** | 表單狀態完整性 - 提交時的狀態應與顯示一致 | Sync 函數只讀取 UI，不重算 |

### 違反這些原則的常見錯誤：
- 在 `SaveDocument` 或 `SyncToModel` 中重新計算金額/稅額
- 用戶手動修改的值被程式覆蓋
- 顯示的數字與資料庫儲存的數字不同
- PostBack 後數值被重算

### 正確做法：
- 計算邏輯只在「值變更事件」中執行，且要保留用戶已修改的值
- `Sync` 函數只負責將 UI 值同步到 Model，不做額外計算
- 若用戶輸入與系統計算不同，顯示警告但**保留用戶的值**

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

## 修正/功能變更規則 (重要!)

當修正錯誤或調整功能時，**必須先詢問用戶**是否需要檢查專案中其他相似的程式碼。

### 工作流程：

1. **先詢問**：「是否需要檢查專案中其他地方有沒有相同或相似的邏輯需要一併修正？」
2. **若用戶說「要」**：
   - 使用 Grep 搜尋相同的函數名稱、變數名稱或邏輯模式
   - 列出所有相關位置給用戶確認
   - 確認修改範圍後再實作
3. **若用戶說「不用」**：
   - 直接修正用戶指定的位置即可

### 注意事項：
- 不要自動假設只改一處或全部都改
- 讓用戶決定修改範圍
- 使用 TodoWrite 追蹤所有需要修改的位置，避免遺漏

## 版號管理與自動 Commit

### 版號格式：X.Y.Z
```
X.Y.Z
```
- **X**: 重大改版（架構調整、不相容變更）
- **Y**: 新功能
- **Z**: 次要更新、日常維護、錯誤修正（**預設**）

**進位規則**：超過 9 直接進位，如 1.0.9 → 1.0.10

### 版號檔案
版號存放於專案根目錄的 `VERSION` 檔案中。

### 自動 Commit 工作流程

當完成用戶要求的修改後（一個要求可能變更多個檔案），**必須主動詢問用戶**是否要 Commit：

1. **分析變更**：執行 `git status` 和 `git diff --stat` 了解變更範圍
2. **產生摘要**：根據變更內容自動產生中文 Commit 訊息
3. **提供版號選項**：使用 AskUserQuestion 工具詢問用戶

### Commit 選項格式

```
請選擇版號更新方式：

目前版號: X.Y.Z

1. [預設] PATCH (X.Y.Z+1) - 錯誤修正、小幅調整
2. MINOR (X.Y+1.0) - 新功能
3. MAJOR (X+1.0.0) - 重大變更
4. 不要 Commit
```

### Commit 訊息格式
```
[vX.Y.Z] 摘要內容

- 變更項目1
- 變更項目2

🤖 Generated with Claude Code
```

### 注意事項
- 每個「用戶要求」完成後才 Commit，不是每個檔案變更都 Commit
- Commit 前必須更新 VERSION 檔案
- 自動產生 git tag: `vX.Y.Z`
- 不自動 push，除非用戶明確要求
