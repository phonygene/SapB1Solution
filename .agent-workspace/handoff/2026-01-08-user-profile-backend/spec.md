# 任務規格：User 表欄位擴充與資料存取模組

> 任務ID：2026-01-08-user-profile-backend
> 指派：Backend Agent
> 優先級：High
> 依賴：無（可先行）

---

## 目標

1. 在 `[User]` 表新增 `EmpSeries`（工號）欄位
2. 建立共用模組 `UserProfileHelper.vb` 提供使用者必填欄位檢查與帳號設定更新功能

---

## 需求詳情

### 1. DDL：新增 EmpSeries 欄位

```sql
ALTER TABLE [User] ADD EmpSeries NVARCHAR(20) NULL;
```

- 欄位名稱：`EmpSeries`
- 類型：`NVARCHAR(20)`
- 允許 NULL（舊帳號可能沒有填）

### 2. 共用模組：UserProfileHelper.vb

位置：`MgmSP/Modules/UserProfileHelper.vb`

#### 2.1 必填欄位檢查

```vb
''' <summary>
''' 檢查使用者必填欄位是否完整
''' </summary>
''' <param name="userId">使用者ID</param>
''' <returns>缺少的欄位列表（空表示全部已填）</returns>
Public Shared Function GetMissingRequiredFields(userId As String) As List(Of String)
```

必填欄位清單：
- `expDEPT`（費用部門）
- `EmpSeries`（工號）

#### 2.2 取得使用者設定資料

```vb
''' <summary>
''' 取得使用者帳號設定資料
''' </summary>
Public Shared Function GetUserProfile(userId As String) As UserProfileData

Public Class UserProfileData
    Public Property UserId As String
    Public Property UserName As String
    Public Property Email As String
    Public Property ExpDept As String
    Public Property EmpSeries As String
End Class
```

#### 2.3 更新使用者設定

```vb
''' <summary>
''' 更新使用者密碼
''' </summary>
Public Shared Function UpdatePassword(userId As String, newPassword As String) As Boolean

''' <summary>
''' 更新使用者費用部門
''' </summary>
Public Shared Function UpdateExpDept(userId As String, expDept As String) As Boolean

''' <summary>
''' 更新使用者工號
''' </summary>
Public Shared Function UpdateEmpSeries(userId As String, empSeries As String) As Boolean

''' <summary>
''' 更新使用者 Email
''' </summary>
Public Shared Function UpdateEmail(userId As String, email As String) As Boolean

''' <summary>
''' 批次更新必填欄位（用於必填欄位彈窗確認）
''' </summary>
Public Shared Function UpdateRequiredFields(userId As String, expDept As String, empSeries As String) As Boolean
```

#### 2.4 載入費用部門清單

```vb
''' <summary>
''' 取得費用部門選項清單
''' </summary>
Public Shared Function GetExpDeptList() As List(Of KeyValuePair(Of String, String))
```

---

## 驗收標準

1. [ ] DDL 執行成功，User 表已有 EmpSeries 欄位
2. [ ] UserProfileHelper.vb 已建立並編譯通過
3. [ ] 所有 Public 方法皆有 XML 註解
4. [ ] 使用參數化查詢防止 SQL Injection

---

## 注意事項

- 連線字串使用 `jtdbConnectionString`
- 密碼儲存方式維持與現有系統一致（明文或 hash，請查看現有邏輯）
- 此模組供 UI-UX Agent 的 Home.aspx.vb 調用

---

## 完成後

1. 在此目錄建立 `output.md` 說明完成項目
2. 通知 Manager Agent 進行審查
