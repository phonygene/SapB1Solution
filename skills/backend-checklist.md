# Backend 開發檢查清單

> 🔗 **核心規則已移至**：`claude-config/core/financial-rules.yaml`
> 本檔案僅保留專案特定的參考資料和錯誤案例

---

## PostBack 處理（專案特定）

- Page_Load 中初始化資料必須包在 `If Not IsPostBack Then` 區塊內
- ViewState 不可存放 ListItem，改用 DataTable
- UpdatePanel 事件必須在 Triggers 註冊 AsyncPostBackTrigger

---

## 常見問題模式

### [P001] COM 物件未釋放
- **症狀**：記憶體增長、SAP 連線過多
- **解法**：`Marshal.ReleaseComObject` + 設為 Nothing
- **觸發規則**：SAP-001

### [P002] PostBack 後資料遺失
- **症狀**：頁面刷新後輸入消失
- **解法**：檢查 `IsPostBack`，使用 ViewState/Session

### [P003] UpdatePanel 事件不觸發
- **症狀**：按鈕點擊無反應
- **解法**：在 Triggers 註冊 AsyncPostBackTrigger

### [P004] ViewState 序列化失敗 (ListItem)
- **症狀**：`SerializationException: 未將類型 'ListItem' 標記為可序列化`
- **解法**：改用 DataTable
```vb
' 錯誤
ViewState("Items") = New List(Of ListItem)()

' 正確
Dim dt As New DataTable()
ViewState("Items") = dt
```

---

## 從錯誤中學習

### 2026-01-08: ViewState 序列化問題
- **檔案**：PurchaseRequestForm.aspx.vb
- **問題**：LoadWarehouses/LoadCostingCodes 使用 `List(Of ListItem)` 存入 ViewState
- **修正**：改用 DataTable，在 RowDataBound 中動態建立 ListItem
