<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ExpenseClaimForm.aspx.vb"
    Inherits="shopfloor.ExpenseClaimForm" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>費用申請單</title>
    <style type="text/css">
        body { font-family: "Microsoft JhengHei", Arial, sans-serif; font-size: 14px; }
        .form-container { max-width: 1400px; margin: 20px auto; padding: 20px; }
        .section-title { background-color: #4CAF50; color: white; padding: 10px;
                        margin: 20px 0 10px 0; font-weight: bold; border-radius: 4px; }
        .form-row { display: flex; margin-bottom: 10px; align-items: center; }
        .form-label { width: 120px; font-weight: bold; padding-right: 10px; text-align: right; }
        .form-field { flex: 1; }
        .form-field input[type="text"], .form-field select, .form-field textarea {
            width: 95%; padding: 5px; border: 1px solid #ccc; border-radius: 3px;
        }
        .form-field textarea { height: 60px; resize: vertical; }
        .form-field-readonly { background-color: #f0f0f0; }
        .button-bar { text-align: center; margin: 20px 0; }
        .btn { padding: 10px 30px; margin: 0 5px; border-radius: 4px;
               border: none; cursor: pointer; font-size: 14px; }
        .btn-primary { background-color: #4CAF50; color: white; }
        .btn-secondary { background-color: #008CBA; color: white; }
        .btn-danger { background-color: #f44336; color: white; }
        .btn-warning { background-color: #ff9800; color: white; }
        .btn:hover { opacity: 0.8; }
        .required { color: red; }
        .tab-container { display: flex; border-bottom: 2px solid #4CAF50; margin: 20px 0; }
        .tab-button { padding: 10px 30px; background-color: #e0e0e0; border: none;
                     cursor: pointer; margin-right: 5px; border-radius: 4px 4px 0 0; }
        .tab-button.active { background-color: #4CAF50; color: white; font-weight: bold; }
        .tab-content { display: none; padding: 20px; border: 1px solid #ccc;
                      border-top: none; border-radius: 0 0 4px 4px; }
        .tab-content.active { display: block; }
        .approval-section { background-color: #fff3cd; padding: 15px;
                           border: 2px solid #ffc107; border-radius: 4px; margin: 20px 0; }
        .status-pending { color: #ff9800; font-weight: bold; }
        .status-approved { color: #4CAF50; font-weight: bold; }
        .status-rejected { color: #f44336; font-weight: bold; }
    </style>
    <script type="text/javascript">
        function switchTab(tabName) {
            // 隱藏所有 tab content
            var contents = document.getElementsByClassName('tab-content');
            for (var i = 0; i < contents.length; i++) {
                contents[i].classList.remove('active');
            }

            // 移除所有 tab button 的 active class
            var buttons = document.getElementsByClassName('tab-button');
            for (var i = 0; i < buttons.length; i++) {
                buttons[i].classList.remove('active');
            }

            // 顯示選中的 tab
            document.getElementById(tabName + '-content').classList.add('active');
            document.getElementById(tabName + '-btn').classList.add('active');
        }

        function confirmDelete() {
            return confirm('確定要刪除此筆費用申請單嗎？');
        }

        function validateForm() {
            // 基本驗證邏輯
            var cardCode = document.getElementById('<%= ddlCardCode.ClientID %>').value;
            if (!cardCode) {
                alert('請選擇供應商');
                return false;
            }
            return true;
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <div class="form-container">
            <!-- 標題與單據資訊 -->
            <h2 style="text-align: center; color: #4CAF50;">費用申請單</h2>

            <div class="form-row">
                <div class="form-label">單據編號:</div>
                <div class="form-field">
                    <asp:Label ID="lblDocNum" runat="server" Text="[自動產生]"
                              Font-Bold="True" ForeColor="Blue"></asp:Label>
                </div>
                <div class="form-label">單據狀態:</div>
                <div class="form-field">
                    <asp:Label ID="lblDocStatus" runat="server" Text="草稿"
                              CssClass="status-pending"></asp:Label>
                </div>
            </div>

            <div class="form-row">
                <div class="form-label">建立人員:</div>
                <div class="form-field">
                    <asp:Label ID="lblCreateBy" runat="server" Text=""></asp:Label>
                </div>
                <div class="form-label">建立日期:</div>
                <div class="form-field">
                    <asp:Label ID="lblCreateDate" runat="server" Text=""></asp:Label>
                </div>
            </div>

            <!-- Tab 切換 -->
            <div class="tab-container">
                <button type="button" class="tab-button active" id="expense-btn"
                        onclick="switchTab('expense')">費用申請</button>
                <button type="button" class="tab-button" id="mdr-btn"
                        onclick="switchTab('mdr')">MDR 發票明細</button>
            </div>

            <!-- 費用申請 Tab -->
            <div id="expense-content" class="tab-content active">
                <div class="section-title">供應商資訊</div>

                <div class="form-row">
                    <div class="form-label"><span class="required">*</span> 供應商代碼:</div>
                    <div class="form-field">
                        <asp:DropDownList ID="ddlCardCode" runat="server" AutoPostBack="True"
                                         OnSelectedIndexChanged="ddlCardCode_SelectedIndexChanged">
                        </asp:DropDownList>
                    </div>
                </div>

                <div class="form-row">
                    <div class="form-label">供應商名稱:</div>
                    <div class="form-field">
                        <asp:TextBox ID="txtCardName" runat="server" ReadOnly="True"
                                    CssClass="form-field-readonly"></asp:TextBox>
                    </div>
                </div>

                <div class="form-row">
                    <div class="form-label">聯絡人:</div>
                    <div class="form-field">
                        <asp:DropDownList ID="ddlContactPerson" runat="server">
                        </asp:DropDownList>
                    </div>
                </div>

                <div class="section-title">單據資訊</div>

                <div class="form-row">
                    <div class="form-label"><span class="required">*</span> 收貨地址:</div>
                    <div class="form-field">
                        <asp:DropDownList ID="ddlDeliveryAddr" runat="server">
                        </asp:DropDownList>
                    </div>
                </div>

                <div class="form-row">
                    <div class="form-label"><span class="required">*</span> 請款日期:</div>
                    <div class="form-field">
                        <asp:TextBox ID="txtDocDate" runat="server" TextMode="Date"></asp:TextBox>
                    </div>
                </div>

                <div class="form-row">
                    <div class="form-label"><span class="required">*</span> 到期日:</div>
                    <div class="form-field">
                        <asp:TextBox ID="txtDocDueDate" runat="server" TextMode="Date"></asp:TextBox>
                    </div>
                </div>

                <div class="form-row">
                    <div class="form-label">產品/部門:</div>
                    <div class="form-field">
                        <asp:DropDownList ID="ddlOcrCode" runat="server">
                        </asp:DropDownList>
                    </div>
                </div>

                <div class="form-row">
                    <div class="form-label">文件幣別:</div>
                    <div class="form-field">
                        <asp:DropDownList ID="ddlDocCurrency" runat="server" AutoPostBack="True"
                                         OnSelectedIndexChanged="ddlDocCurrency_SelectedIndexChanged">
                        </asp:DropDownList>
                    </div>
                </div>

                <div class="form-row">
                    <div class="form-label">匯率:</div>
                    <div class="form-field">
                        <asp:TextBox ID="txtDocRate" runat="server" Text="1.0"></asp:TextBox>
                        <asp:Button ID="btnRefreshRate" runat="server" Text="更新匯率"
                                   OnClick="btnRefreshRate_Click" CssClass="btn btn-secondary"
                                   Style="width: auto; padding: 5px 15px; margin-left: 10px;" />
                    </div>
                </div>

                <div class="form-row">
                    <div class="form-label">備註:</div>
                    <div class="form-field">
                        <asp:TextBox ID="txtRemarks" runat="server" TextMode="MultiLine"></asp:TextBox>
                    </div>
                </div>

                <div class="section-title">附件上傳</div>

                <div class="form-row">
                    <div class="form-label">選擇檔案:</div>
                    <div class="form-field">
                        <asp:FileUpload ID="fileUpload" runat="server" />
                        <asp:Button ID="btnUpload" runat="server" Text="上傳"
                                   OnClick="btnUpload_Click" CssClass="btn btn-secondary"
                                   Style="width: auto; padding: 5px 15px; margin-left: 10px;" />
                    </div>
                </div>

                <div class="form-row">
                    <div class="form-label">已上傳檔案:</div>
                    <div class="form-field">
                        <asp:Label ID="lblAttachment" runat="server" Text="無"></asp:Label>
                        <asp:Button ID="btnDownload" runat="server" Text="下載"
                                   OnClick="btnDownload_Click" CssClass="btn btn-secondary"
                                   Style="width: auto; padding: 5px 15px; margin-left: 10px;"
                                   Visible="False" />
                    </div>
                </div>

                <!-- 審核區塊（僅在有審核權限時顯示） -->
                <asp:Panel ID="pnlApproval" runat="server" CssClass="approval-section" Visible="False">
                    <div class="section-title" style="background-color: #ffc107;">審核資訊</div>

                    <div class="form-row">
                        <div class="form-label">審核狀態:</div>
                        <div class="form-field">
                            <asp:Label ID="lblApprovalStatus" runat="server" Text="待審核"
                                      CssClass="status-pending"></asp:Label>
                        </div>
                    </div>

                    <div class="form-row">
                        <div class="form-label">審核人員:</div>
                        <div class="form-field">
                            <asp:Label ID="lblApprovedBy" runat="server" Text=""></asp:Label>
                        </div>
                        <div class="form-label">審核日期:</div>
                        <div class="form-field">
                            <asp:Label ID="lblApprovedDate" runat="server" Text=""></asp:Label>
                        </div>
                    </div>

                    <div class="form-row">
                        <div class="form-label">審核意見:</div>
                        <div class="form-field">
                            <asp:TextBox ID="txtApprovalComments" runat="server" TextMode="MultiLine"
                                        Height="80px"></asp:TextBox>
                        </div>
                    </div>

                    <div class="form-row" style="justify-content: center;">
                        <asp:Button ID="btnApprove" runat="server" Text="放行"
                                   OnClick="btnApprove_Click" CssClass="btn btn-primary" />
                        <asp:Button ID="btnReject" runat="server" Text="駁回"
                                   OnClick="btnReject_Click" CssClass="btn btn-danger" />
                        <asp:Button ID="btnSendNotification" runat="server" Text="發送通知"
                                   OnClick="btnSendNotification_Click" CssClass="btn btn-warning" />
                    </div>
                </asp:Panel>

                <!-- 費用明細 GridView - 下一階段實作 -->
                <div class="section-title">費用明細</div>
                <div>
                    <p>[明細 GridView 將在第二階段實作]</p>
                </div>
            </div>

            <!-- MDR Tab - 下一階段實作 -->
            <div id="mdr-content" class="tab-content">
                <div class="section-title">MDR 發票明細（唯讀，自動同步）</div>
                <div>
                    <p>[MDR Tab 內容將在第二階段實作]</p>
                </div>
            </div>

            <!-- 按鈕列 -->
            <div class="button-bar">
                <asp:Button ID="btnSave" runat="server" Text="儲存"
                           OnClick="btnSave_Click" CssClass="btn btn-primary" />
                <asp:Button ID="btnSubmit" runat="server" Text="送出"
                           OnClick="btnSubmit_Click" CssClass="btn btn-primary"
                           OnClientClick="return validateForm();" />
                <asp:Button ID="btnDelete" runat="server" Text="刪除"
                           OnClick="btnDelete_Click" CssClass="btn btn-danger"
                           OnClientClick="return confirmDelete();" />
                <asp:Button ID="btnCancel" runat="server" Text="取消"
                           OnClick="btnCancel_Click" CssClass="btn btn-secondary" />
            </div>

            <!-- 訊息顯示 -->
            <div style="text-align: center; margin-top: 20px;">
                <asp:Label ID="lblMessage" runat="server" Text=""
                          Font-Bold="True" ForeColor="Red"></asp:Label>
            </div>
        </div>
    </form>
</body>
</html>
