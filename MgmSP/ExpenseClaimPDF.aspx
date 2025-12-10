<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ExpenseClaimPDF.aspx.vb" Inherits="ExpenseClaimPDF" %>
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <title>費用申請單</title>
    <style type="text/css">
        @media print {
            .no-print { display: none !important; }
            body { margin: 0; padding: 10mm; }
        }
        body {
            font-family: "Microsoft JhengHei", "微軟正黑體", Arial, sans-serif;
            font-size: 12px;
            line-height: 1.4;
            margin: 20px;
        }
        .header {
            text-align: center;
            margin-bottom: 20px;
        }
        .header h1 {
            font-size: 24px;
            margin: 0 0 10px 0;
        }
        .header .doc-info {
            font-size: 14px;
            color: #666;
        }
        .section {
            margin-bottom: 15px;
        }
        .section-title {
            font-size: 14px;
            font-weight: bold;
            background-color: #f0f0f0;
            padding: 5px 10px;
            border-left: 4px solid #333;
            margin-bottom: 10px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 15px;
        }
        table th, table td {
            border: 1px solid #ccc;
            padding: 6px 8px;
            text-align: left;
        }
        table th {
            background-color: #f5f5f5;
            font-weight: bold;
            width: 120px;
        }
        .detail-table th {
            text-align: center;
            width: auto;
        }
        .detail-table td {
            text-align: center;
        }
        .detail-table td.text-left {
            text-align: left;
        }
        .detail-table td.text-right {
            text-align: right;
        }
        .total-row {
            font-weight: bold;
            background-color: #f9f9f9;
        }
        .amount {
            text-align: right;
            font-family: Consolas, monospace;
        }
        .footer {
            margin-top: 30px;
            border-top: 1px solid #ccc;
            padding-top: 15px;
        }
        .signature-area {
            display: flex;
            justify-content: space-between;
            margin-top: 40px;
        }
        .signature-box {
            width: 30%;
            text-align: center;
        }
        .signature-line {
            border-top: 1px solid #333;
            margin-top: 50px;
            padding-top: 5px;
        }
        .print-button {
            position: fixed;
            top: 20px;
            right: 20px;
            padding: 10px 20px;
            font-size: 14px;
            cursor: pointer;
            background-color: #4CAF50;
            color: white;
            border: none;
            border-radius: 4px;
        }
        .print-button:hover {
            background-color: #45a049;
        }
        .back-button {
            position: fixed;
            top: 20px;
            right: 120px;
            padding: 10px 20px;
            font-size: 14px;
            cursor: pointer;
            background-color: #666;
            color: white;
            border: none;
            border-radius: 4px;
            text-decoration: none;
        }
        .back-button:hover {
            background-color: #555;
        }
        .status-badge {
            display: inline-block;
            padding: 2px 8px;
            border-radius: 3px;
            font-size: 11px;
            font-weight: bold;
        }
        .status-W { background-color: #fff3cd; color: #856404; }
        .status-A { background-color: #d4edda; color: #155724; }
        .status-R { background-color: #f8d7da; color: #721c24; }
        .status-P { background-color: #e2e3e5; color: #383d41; }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <!-- 列印按鈕 (列印時隱藏) -->
        <a href="ExpenseClaimForm.aspx?DocEntry=<%= Request.QueryString("DocEntry") %>" class="back-button no-print">返回</a>
        <button type="button" onclick="window.print();" class="print-button no-print">列印 / 另存 PDF</button>

        <!-- 標題 -->
        <div class="header">
            <h1>費用申請單</h1>
            <div class="doc-info">
                單據編號: <asp:Literal ID="litDocNum" runat="server"></asp:Literal>
                &nbsp;&nbsp;|&nbsp;&nbsp;
                狀態: <asp:Literal ID="litStatus" runat="server"></asp:Literal>
            </div>
        </div>

        <!-- 基本資料 -->
        <div class="section">
            <div class="section-title">基本資料</div>
            <table>
                <tr>
                    <th>供應商代碼</th>
                    <td><asp:Literal ID="litCardCode" runat="server"></asp:Literal></td>
                    <th>供應商名稱</th>
                    <td><asp:Literal ID="litCardName" runat="server"></asp:Literal></td>
                </tr>
                <tr>
                    <th>供應商參考號</th>
                    <td><asp:Literal ID="litNumAtCard" runat="server"></asp:Literal></td>
                    <th>採購人員</th>
                    <td><asp:Literal ID="litPurchaser" runat="server"></asp:Literal></td>
                </tr>
                <tr>
                    <th>過帳日期</th>
                    <td><asp:Literal ID="litDocDate" runat="server"></asp:Literal></td>
                    <th>到期日</th>
                    <td><asp:Literal ID="litDocDueDate" runat="server"></asp:Literal></td>
                </tr>
                <tr>
                    <th>文件日期</th>
                    <td><asp:Literal ID="litTaxDate" runat="server"></asp:Literal></td>
                    <th>幣別 / 匯率</th>
                    <td><asp:Literal ID="litCurrency" runat="server"></asp:Literal></td>
                </tr>
                <tr>
                    <th>付款條件</th>
                    <td colspan="3"><asp:Literal ID="litPaymentTerms" runat="server"></asp:Literal></td>
                </tr>
                <tr>
                    <th>備註</th>
                    <td colspan="3"><asp:Literal ID="litRemarks" runat="server"></asp:Literal></td>
                </tr>
            </table>
        </div>

        <!-- 費用明細 -->
        <div class="section">
            <div class="section-title">費用明細</div>
            <asp:Repeater ID="rptExpenseLines" runat="server">
                <HeaderTemplate>
                    <table class="detail-table">
                        <tr>
                            <th style="width:40px;">#</th>
                            <th style="width:100px;">費用項目</th>
                            <th>說明</th>
                            <th style="width:80px;">科目代碼</th>
                            <th style="width:80px;">稅別</th>
                            <th style="width:100px;">未稅金額</th>
                            <th style="width:80px;">稅額</th>
                            <th style="width:100px;">含稅金額</th>
                        </tr>
                </HeaderTemplate>
                <ItemTemplate>
                    <tr>
                        <td><%# Eval("LineNum") %></td>
                        <td class="text-left"><%# Eval("CategoryName") %></td>
                        <td class="text-left"><%# Eval("Description") %></td>
                        <td><%# Eval("AcctCode") %></td>
                        <td><%# Eval("VatGroupName") %></td>
                        <td class="text-right"><%# Eval("LineTotal", "{0:N0}") %></td>
                        <td class="text-right"><%# Eval("LineVat", "{0:N0}") %></td>
                        <td class="text-right"><%# Eval("GTotal", "{0:N0}") %></td>
                    </tr>
                </ItemTemplate>
                <FooterTemplate>
                    </table>
                </FooterTemplate>
            </asp:Repeater>
        </div>

        <!-- 憑證明細 -->
        <asp:Panel ID="pnlMDR" runat="server" Visible="false">
            <div class="section">
                <div class="section-title">憑證明細 (營業稅申報)</div>
                <asp:Repeater ID="rptMDRLines" runat="server">
                    <HeaderTemplate>
                        <table class="detail-table">
                            <tr>
                                <th style="width:40px;">#</th>
                                <th style="width:100px;">統一編號</th>
                                <th style="width:120px;">憑證號碼</th>
                                <th style="width:100px;">憑證類型</th>
                                <th style="width:90px;">憑證日期</th>
                                <th style="width:100px;">未稅金額</th>
                                <th style="width:80px;">稅額</th>
                            </tr>
                    </HeaderTemplate>
                    <ItemTemplate>
                        <tr>
                            <td><%# Eval("LineNum") %></td>
                            <td><%# Eval("U_STCEG") %></td>
                            <td><%# Eval("U_XBLNR") %></td>
                            <td><%# Eval("ZFormName") %></td>
                            <td><%# Eval("U_BLDAT", "{0:yyyy-MM-dd}") %></td>
                            <td class="text-right"><%# Eval("U_HWBAS", "{0:N0}") %></td>
                            <td class="text-right"><%# Eval("U_HWSTE", "{0:N0}") %></td>
                        </tr>
                    </ItemTemplate>
                    <FooterTemplate>
                        </table>
                    </FooterTemplate>
                </asp:Repeater>
            </div>
        </asp:Panel>

        <!-- 金額彙總 -->
        <div class="section">
            <div class="section-title">金額彙總</div>
            <table style="width:400px; margin-left:auto;">
                <tr>
                    <th>未稅總額</th>
                    <td class="amount"><asp:Literal ID="litDocTotal" runat="server"></asp:Literal></td>
                </tr>
                <tr>
                    <th>稅額合計</th>
                    <td class="amount"><asp:Literal ID="litVatSum" runat="server"></asp:Literal></td>
                </tr>
                <tr class="total-row">
                    <th>含稅總額</th>
                    <td class="amount"><asp:Literal ID="litGrandTotal" runat="server"></asp:Literal></td>
                </tr>
            </table>
        </div>

        <!-- 簽核資訊 -->
        <div class="footer">
            <div class="section-title">簽核資訊</div>
            <table>
                <tr>
                    <th>申請人</th>
                    <td><asp:Literal ID="litCreateBy" runat="server"></asp:Literal></td>
                    <th>申請日期</th>
                    <td><asp:Literal ID="litCreateDate" runat="server"></asp:Literal></td>
                </tr>
                <tr>
                    <th>核准人</th>
                    <td><asp:Literal ID="litApprovedBy" runat="server"></asp:Literal></td>
                    <th>核准日期</th>
                    <td><asp:Literal ID="litApprovalDate" runat="server"></asp:Literal></td>
                </tr>
                <tr>
                    <th>審核意見</th>
                    <td colspan="3"><asp:Literal ID="litApprovalComments" runat="server"></asp:Literal></td>
                </tr>
            </table>
        </div>

        <!-- 簽名欄 (列印用) -->
        <div class="signature-area">
            <div class="signature-box">
                <div class="signature-line">申請人</div>
            </div>
            <div class="signature-box">
                <div class="signature-line">部門主管</div>
            </div>
            <div class="signature-box">
                <div class="signature-line">財務核准</div>
            </div>
        </div>

    </form>
</body>
</html>
