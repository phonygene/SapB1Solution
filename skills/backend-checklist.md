# Backend 開發檢查清單

> Backend Agent 專用

---

## 資料庫操作

- [ ] SQL 使用參數化查詢，禁止字串拼接
- [ ] 使用 `Using` 確保連線釋放
- [ ] 多表操作使用 Transaction

## SAP B1 整合

- [ ] Service Layer 呼叫前檢查 Session
- [ ] COM 物件使用後必須釋放 `Marshal.ReleaseComObject`
- [ ] 日期格式 yyyy-MM-dd
- [ ] 有處理 SAP 錯誤碼 `GetLastErrorCode()`

## 金額處理

- [ ] 金額使用 `Decimal`，不用 `Double`
- [ ] 四捨五入使用 `Math.Round(..., MidpointRounding.AwayFromZero)`

## 計算邏輯（財務系統核心）

- [ ] 計算只在「值變更事件」中執行
- [ ] **不在 Save 時重新計算用戶輸入**
- [ ] Sync 函數只讀取 UI，不重算
- [ ] 用戶手動修改的值優先保留

## PostBack 處理

- [ ] 正確使用 `IsPostBack` 判斷
- [ ] ViewState 敏感欄位有處理
- [ ] UpdatePanel 事件有在 Triggers 註冊

---

## 常見問題模式

### [P001] COM 物件未釋放
- **症狀**：記憶體增長、SAP 連線過多
- **解法**：`Marshal.ReleaseComObject` + 設為 Nothing

### [P002] PostBack 後資料遺失
- **症狀**：頁面刷新後輸入消失
- **解法**：檢查 `IsPostBack`，使用 ViewState/Session

### [P003] UpdatePanel 事件不觸發
- **症狀**：按鈕點擊無反應
- **解法**：在 Triggers 註冊 AsyncPostBackTrigger

---

## 從錯誤中學習（持續新增）

*（待累積）*
