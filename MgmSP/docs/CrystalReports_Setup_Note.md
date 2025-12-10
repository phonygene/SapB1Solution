# Crystal Reports 參考設定說明

## 問題背景
專案原本參考的 Crystal Reports DLL 版本是 13.0.4000.0（SAP BusinessObjects 完整安裝版），但系統安裝的 32-bit Runtime 是 13.0.2000.0 版本，導致版本不相容錯誤。

## 解決方案
將專案參考改為使用 13.0.2000.0 版本（配合 32-bit Runtime）。

---

## 步驟 1：移除現有的 Crystal Reports 參考

1. 在 Visual Studio 中開啟專案
2. 展開「方案總管」中的「參考」節點
3. 找到以下 Crystal Reports 相關參考並**逐一移除**（右鍵 > 移除）：
   - `CrystalDecisions.CrystalReports.Engine`
   - `CrystalDecisions.Shared`
   - `CrystalDecisions.ReportAppServer.ClientDoc`
   - `CrystalDecisions.ReportAppServer.CommLayer`
   - `CrystalDecisions.ReportAppServer.DataDefModel`
   - `CrystalDecisions.ReportAppServer.ReportDefModel`
   - `CrystalDecisions.ReportSource`

---

## 步驟 2：重新加入參考

### 方法 A：從 GAC 自動解析（建議）

1. 右鍵點擊「參考」>「加入參考」
2. 選擇「組件」>「延伸模組」
3. 在清單中搜尋 `CrystalDecisions`
4. 勾選以下組件（確認版本為 13.0.2000.0）：
   - `CrystalDecisions.CrystalReports.Engine`
   - `CrystalDecisions.Shared`
   - `CrystalDecisions.ReportAppServer.ClientDoc`
   - `CrystalDecisions.ReportAppServer.CommLayer`
   - `CrystalDecisions.ReportAppServer.DataDefModel`
   - `CrystalDecisions.ReportAppServer.ReportDefModel`
   - `CrystalDecisions.ReportSource`
5. 按「確定」完成

### 方法 B：手動從路徑加入

如果方法 A 找不到組件，可從以下路徑手動加入：

**GAC 路徑（32-bit）：**
```
C:\Windows\Microsoft.NET\assembly\GAC_32\
```

各組件的完整路徑：
- `C:\Windows\Microsoft.NET\assembly\GAC_32\CrystalDecisions.CrystalReports.Engine\v4.0_13.0.2000.0__692fbea5521e1304\CrystalDecisions.CrystalReports.Engine.dll`
- `C:\Windows\Microsoft.NET\assembly\GAC_MSIL\CrystalDecisions.Shared\v4.0_13.0.2000.0__692fbea5521e1304\CrystalDecisions.Shared.dll`

**或從 Runtime 安裝目錄：**
```
C:\Program Files (x86)\SAP BusinessObjects\Crystal Reports for .NET Framework 4.0\Common\SAP BusinessObjects Enterprise XI 4.0\win32_x86\
```

操作步驟：
1. 右鍵點擊「參考」>「加入參考」
2. 選擇「瀏覽」
3. 導航到上述路徑
4. 選擇需要的 DLL 檔案
5. 按「加入」>「確定」

---

## 步驟 3：驗證參考設定

加入參考後，檢查每個 Crystal Reports 參考的屬性：

1. 點擊參考項目
2. 在「屬性」視窗確認：
   - **版本**：13.0.2000.0
   - **複製到本機 (Copy Local)**：False
   - **特定版本 (Specific Version)**：False

---

## 步驟 4：重新編譯與測試

1. 清除方案：「建置」>「清除方案」
2. 重建方案：「建置」>「重建方案」
3. 確認編譯成功無錯誤
4. 執行應用程式
5. 測試費用申請單的「匯出PDF」功能

---

## 疑難排解

### 如果仍出現版本錯誤
確認 Web.config 中的 binding redirect 設定：
```xml
<dependentAssembly>
  <assemblyIdentity name="CrystalDecisions.Shared" publicKeyToken="692fbea5521e1304" culture="neutral" />
  <bindingRedirect oldVersion="0.0.0.0-13.0.9999.0" newVersion="13.0.2000.0" />
</dependentAssembly>
```

### 如果找不到組件
確認已安裝 Crystal Reports Runtime 32-bit：
- 下載：SAP Crystal Reports runtime engine for .NET Framework (32-bit)
- 安裝後重新開啟 Visual Studio

### 如果 IIS 應用程式集區錯誤
確認應用程式集區設定：
1. 開啟 IIS 管理員
2. 找到應用程式集區
3. 進階設定 > 啟用 32 位元應用程式 = True

---

## 相關檔案
- `MgmSP.vbproj` - 專案參考設定
- `Web.config` - 組件繫結重定向
- `ExpenseClaimReport.ashx.vb` - PDF 匯出程式碼
