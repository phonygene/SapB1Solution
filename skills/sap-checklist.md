# SAP B1 整合檢查清單

> 涉及 SAP Business One 整合時使用

---

## Service Layer

- [ ] 呼叫前檢查 Session 是否有效
- [ ] 處理 Session 過期的情況
- [ ] 正確處理 HTTP 錯誤碼

## DI API

- [ ] COM 物件使用後釋放
  ```vb
  Marshal.ReleaseComObject(oDoc)
  oDoc = Nothing
  ```
- [ ] 檢查 `GetLastErrorCode()`
- [ ] 記錄 `GetLastErrorDescription()`

## 資料格式

- [ ] 日期格式：`yyyy-MM-dd`
- [ ] 金額使用 `Decimal`
- [ ] 字串長度符合 SAP 欄位限制

## 欄位對應

使用 `[AI-Context]` 註解標記 SAP 欄位對應：

```vb
' [AI-Context] SAP Table: OEXD (費用申請主表), 欄位: DocTotal
Dim totalAmount As Decimal = ...

' [AI-Context] SAP Table: EXD1 (費用申請明細), 欄位: LineTotal
Dim lineAmount As Decimal = ...
```

## 常用 SAP 表格

| 表格 | 說明 |
|------|------|
| OITM | 物料主檔 |
| OCRD | 業務夥伴主檔 |
| OPCH | 採購發票主表 |
| PCH1 | 採購發票明細 |
| OINV | 銷售發票主表 |
| INV1 | 銷售發票明細 |
| ORDR | 銷售訂單主表 |
| RDR1 | 銷售訂單明細 |

---

## 從錯誤中學習（持續新增）

*（待累積）*
