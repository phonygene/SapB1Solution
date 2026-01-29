# SAP B1 整合參考

> 🔗 **核心規則已移至**：`claude-config/core/financial-rules.yaml`
> 本檔案僅保留 SAP 特定的參考資料

---

## [AI-Context] 註解範例

使用 `[AI-Context]` 標記 SAP 欄位對應：

```vb
' [AI-Context] SAP Table: OEXD (費用申請主表), 欄位: DocTotal
Dim totalAmount As Decimal = ...

' [AI-Context] SAP Table: EXD1 (費用申請明細), 欄位: LineTotal
Dim lineAmount As Decimal = ...
```

---

## 常用 SAP 表格參考

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
| OPRQ | 採購申請主表 |
| PRQ1 | 採購申請明細 |

---

## SAP 欄位長度限制（常見）

| 欄位 | 長度 | 說明 |
|------|------|------|
| ItemCode | 50 | 物料編號 |
| ItemName | 100 | 物料名稱 |
| CardCode | 15 | 業務夥伴代碼 |
| CardName | 100 | 業務夥伴名稱 |
| Comments | 254 | 備註 |

---

## 從錯誤中學習

*（待累積）*
