<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="DocumentSearch.aspx.vb" Inherits="MgmSP.DocumentSearch"
    MaintainScrollPositionOnPostback="true" MasterPageFile="~/MySite1.Master" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <title>單據查詢</title>
    <style type="text/css">
        /* ============================================
           DocumentSearch - 頁面專用樣式
           共用樣式來自 components.css
           變數來自 jet-color-themes.css
           ============================================ */

        /* 篩選區域輸入控制項 */
        .filter-control input[type="text"],
        .filter-control input[type="date"],
        .filter-control select {
            padding: 8px 12px;
            border: 1px solid var(--border-color);
            border-radius: var(--radius-md);
            font-size: 13px;
            font-family: inherit;
            color: var(--text-primary);
            background-color: var(--bg-white);
            transition: all 0.2s ease;
        }

        .filter-control input[type="text"]:focus,
        .filter-control input[type="date"]:focus,
        .filter-control select:focus {
            outline: none;
            border-color: var(--accent-light);
            box-shadow: 0 0 0 3px rgba(0, 0, 0, 0.06);
        }

        .filter-control input::placeholder {
            color: var(--text-muted);
        }

        /* Radio Button */
        .filter-control input[type="radio"] {
            accent-color: var(--accent-primary);
        }

        .filter-control label {
            color: var(--text-secondary);
            font-size: 12px;
            margin-left: 2px;
            margin-right: 8px;
        }

        /* 狀態標籤 - 對應資料庫 ApprovalStatus 值 */
        /* 設計原則：深色/飽和背景 + 白色文字 */
        .status-P { background-color: #64748B; color: #FFFFFF; }  /* 草稿 - 灰 */
        .status-W { background-color: #D97706; color: #FFFFFF; }  /* 待審 - 琥珀 */
        .status-A { background-color: #059669; color: #FFFFFF; }  /* 核准 - 翠綠 */
        .status-R { background-color: #DC2626; color: #FFFFFF; }  /* 退回 - 紅 */

        /* jID 連結 - 使用主題主色 */
        .link-jid {
            color: var(--accent-primary);
            text-decoration: none;
            font-weight: 600;
            transition: color 0.2s ease;
        }

        .link-jid:hover {
            color: var(--accent-hover);
            text-decoration: underline;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server"></asp:ScriptManager>

    <div class="form-container">
        <h2 class="page-title">單據查詢</h2>

        <div class="section-header">篩選條件</div>

        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
            <ContentTemplate>
                <!-- Row 1 -->
                <div class="filter-row">
                    <div class="filter-group">
                        <span class="filter-label">單據類型:</span>
                        <div class="filter-control">
                            <asp:DropDownList ID="ddlDocType" runat="server" Width="120px" AutoPostBack="true" OnSelectedIndexChanged="ddlDocType_SelectedIndexChanged">
                                <asp:ListItem Value="ExpenseClaim" Text="費用申請單"></asp:ListItem>
                                <asp:ListItem Value="PurchaseRequest" Text="請購單"></asp:ListItem>
                            </asp:DropDownList>
                        </div>
                    </div>
                    <div class="filter-group">
                        <span class="filter-label">使用者代碼:</span>
                        <div class="filter-control">
                            <asp:TextBox ID="txtUserCode" runat="server" Width="80px"></asp:TextBox>
                        </div>
                    </div>
                    <div class="filter-group">
                        <span class="filter-label">使用者名稱:</span>
                        <div class="filter-control">
                            <asp:TextBox ID="txtUserName" runat="server" Width="100px"></asp:TextBox>
                        </div>
                    </div>
                    <div class="filter-group">
                        <span class="filter-label">jID:</span>
                        <div class="filter-control">
                            <asp:TextBox ID="txtJID" runat="server" Width="80px"></asp:TextBox>
                        </div>
                    </div>
                </div>

                <!-- Row 2 -->
                <div class="filter-row">
                    <div class="filter-group">
                        <span class="filter-label">AP單號:</span>
                        <div class="filter-control">
                            <asp:TextBox ID="txtAPNumFrom" runat="server" Width="80px" placeholder="起"></asp:TextBox>
                            <span class="range-sep">~</span>
                            <asp:TextBox ID="txtAPNumTo" runat="server" Width="80px" placeholder="迄"></asp:TextBox>
                        </div>
                    </div>
                    <div class="filter-group">
                        <span class="filter-label">簽核系統PID:</span>
                        <div class="filter-control">
                            <asp:TextBox ID="txtPIDFrom" runat="server" Width="80px" placeholder="起"></asp:TextBox>
                            <span class="range-sep">~</span>
                            <asp:TextBox ID="txtPIDTo" runat="server" Width="80px" placeholder="迄"></asp:TextBox>
                        </div>
                    </div>
                    <div class="filter-group">
                        <span class="filter-label">文件狀態:</span>
                        <div class="filter-control">
                            <asp:DropDownList ID="ddlDocStatus" runat="server" Width="100px">
                                <asp:ListItem Value="" Text="全部"></asp:ListItem>
                                <asp:ListItem Value="P" Text="草稿"></asp:ListItem>
                                <asp:ListItem Value="W" Text="待審核"></asp:ListItem>
                                <asp:ListItem Value="A" Text="已核准"></asp:ListItem>
                                <asp:ListItem Value="R" Text="已退回"></asp:ListItem>
                            </asp:DropDownList>
                        </div>
                    </div>
                    <div class="filter-group">
                        <span class="filter-label">放行狀態:</span>
                        <div class="filter-control">
                            <asp:DropDownList ID="ddlApprovalStatus" runat="server" Width="100px">
                                <asp:ListItem Value="" Text="全部"></asp:ListItem>
                                <asp:ListItem Value="Y" Text="已放行"></asp:ListItem>
                                <asp:ListItem Value="N" Text="未放行"></asp:ListItem>
                            </asp:DropDownList>
                        </div>
                    </div>
                </div>

                <!-- Row 3 -->
                <div class="filter-row">
                    <div class="filter-group">
                        <span class="filter-label">文件日期:</span>
                        <div class="filter-control">
                            <asp:TextBox ID="txtDocDateFrom" runat="server" TextMode="Date" Width="130px"></asp:TextBox>
                            <span class="range-sep">~</span>
                            <asp:TextBox ID="txtDocDateTo" runat="server" TextMode="Date" Width="130px"></asp:TextBox>
                        </div>
                    </div>
                    <div class="filter-group">
                        <span class="filter-label">到期日:</span>
                        <div class="filter-control">
                            <asp:TextBox ID="txtDueDateFrom" runat="server" TextMode="Date" Width="130px"></asp:TextBox>
                            <span class="range-sep">~</span>
                            <asp:TextBox ID="txtDueDateTo" runat="server" TextMode="Date" Width="130px"></asp:TextBox>
                        </div>
                    </div>
                    <div class="filter-group">
                        <span class="filter-label">過帳日期:</span>
                        <div class="filter-control">
                            <asp:TextBox ID="txtTaxDateFrom" runat="server" TextMode="Date" Width="130px"></asp:TextBox>
                            <span class="range-sep">~</span>
                            <asp:TextBox ID="txtTaxDateTo" runat="server" TextMode="Date" Width="130px"></asp:TextBox>
                        </div>
                    </div>
                </div>

                <!-- Row 4 -->
                <div class="filter-row">
                    <div class="filter-group">
                        <span class="filter-label">供應商代碼:</span>
                        <div class="filter-control">
                            <asp:TextBox ID="txtCardCodeFrom" runat="server" Width="80px" placeholder="起"></asp:TextBox>
                            <span class="range-sep">~</span>
                            <asp:TextBox ID="txtCardCodeTo" runat="server" Width="80px" placeholder="迄"></asp:TextBox>
                        </div>
                    </div>
                    <div class="filter-group">
                        <span class="filter-label">供應商名稱:</span>
                        <div class="filter-control">
                            <asp:TextBox ID="txtCardName" runat="server" Width="120px"></asp:TextBox>
                            <asp:RadioButtonList ID="rblCardNameMode" runat="server"
                                RepeatDirection="Horizontal" style="display:inline-block; margin-left:5px;">
                                <asp:ListItem Value="StartsWith" Text="開頭" Selected="True"></asp:ListItem>
                                <asp:ListItem Value="Contains" Text="模糊"></asp:ListItem>
                            </asp:RadioButtonList>
                        </div>
                    </div>
                    <div class="filter-group">
                        <span class="filter-label">備註:</span>
                        <div class="filter-control">
                            <asp:TextBox ID="txtComments" runat="server" Width="120px"></asp:TextBox>
                            <asp:RadioButtonList ID="rblCommentsMode" runat="server"
                                RepeatDirection="Horizontal" style="display:inline-block; margin-left:5px;">
                                <asp:ListItem Value="StartsWith" Text="開頭" Selected="True"></asp:ListItem>
                                <asp:ListItem Value="Contains" Text="模糊"></asp:ListItem>
                            </asp:RadioButtonList>
                        </div>
                    </div>
                </div>

                <!-- Row 5: Sorting & Paging -->
                <div class="filter-row filter-divider">
                    <div class="filter-group">
                        <span class="filter-label">排序依據:</span>
                        <div class="filter-control">
                            <asp:DropDownList ID="ddlSortBy" runat="server" Width="120px">
                                <asp:ListItem Value="jID" Text="jID" Selected="True"></asp:ListItem>
                                <asp:ListItem Value="DocDate" Text="文件日期"></asp:ListItem>
                                <asp:ListItem Value="DocDueDate" Text="到期日"></asp:ListItem>
                                <asp:ListItem Value="CardCode" Text="供應商代碼"></asp:ListItem>
                                <asp:ListItem Value="CardName" Text="供應商名稱"></asp:ListItem>
                                <asp:ListItem Value="CreateDate" Text="建立日期"></asp:ListItem>
                            </asp:DropDownList>
                        </div>
                    </div>
                    <div class="filter-group">
                        <asp:RadioButtonList ID="rblSortOrder" runat="server" RepeatDirection="Horizontal">
                            <asp:ListItem Value="DESC" Text="倒序" Selected="True"></asp:ListItem>
                            <asp:ListItem Value="ASC" Text="正序"></asp:ListItem>
                        </asp:RadioButtonList>
                    </div>
                    <div class="filter-group">
                        <span class="filter-label">每頁筆數:</span>
                        <div class="filter-control">
                            <asp:DropDownList ID="ddlPageSize" runat="server" Width="70px">
                                <asp:ListItem Value="10" Text="10"></asp:ListItem>
                                <asp:ListItem Value="20" Text="20" Selected="True"></asp:ListItem>
                                <asp:ListItem Value="50" Text="50"></asp:ListItem>
                                <asp:ListItem Value="100" Text="100"></asp:ListItem>
                            </asp:DropDownList>
                        </div>
                    </div>
                    <div class="filter-group" style="margin-left:auto;">
                        <asp:Button ID="btnSearch" runat="server" Text="查詢" CssClass="btn btn-primary" OnClick="btnSearch_Click" />
                        <asp:Button ID="btnClear" runat="server" Text="清除條件" CssClass="btn btn-secondary" OnClick="btnClear_Click" />
                    </div>
                </div>

                <!-- Message -->
                <asp:Label ID="lblMessage" runat="server" Font-Bold="True"></asp:Label>

                <!-- Results Section -->
                <div class="section-header">查詢結果</div>
                <asp:Label ID="lblResultCount" runat="server" CssClass="result-count"></asp:Label>

                <asp:HiddenField ID="hfCopyAttachment" runat="server" Value="0" />
                <asp:HiddenField ID="hfCopyMDR" runat="server" Value="0" />
                <asp:HiddenField ID="hfCopyRowIndex" runat="server" Value="" />
                <asp:Button ID="btnCopyConfirm" runat="server" style="display:none" OnClick="btnCopyConfirm_Click" />

                <div id="divCopyModal" class="modal-overlay">
                    <div class="modal-content">
                        <div class="modal-title">複製選項</div>
                        <div id="divCopyQuestion" class="modal-body">是否複製附件？</div>
                        <div class="modal-footer">
                            <button type="button" class="btn btn-primary" onclick="copyDialogAnswer('yes');">是</button>
                            <button type="button" class="btn btn-secondary" onclick="copyDialogAnswer('no');">否</button>
                            <button type="button" class="btn btn-secondary" onclick="copyDialogAnswer('cancel');">取消</button>
                        </div>
                    </div>
                </div>

                <script type="text/javascript">
                    var copyStage = '';
                    function showCopyDialog(rowIndex) {
                        var hfRow = document.getElementById('<%= hfCopyRowIndex.ClientID %>');
                        if (hfRow) {
                            hfRow.value = rowIndex;
                        }
                        copyStage = 'attach';
                        var question = document.getElementById('divCopyQuestion');
                        if (question) {
                            question.innerText = '是否複製附件？';
                        }
                        var modal = document.getElementById('divCopyModal');
                        if (modal) {
                            modal.style.display = 'block';
                        }
                        return false;
                    }

                    function copyDialogAnswer(answer) {
                        if (answer === 'cancel') {
                            hideCopyDialog();
                            return;
                        }
                        if (copyStage === 'attach') {
                            var hfAttach = document.getElementById('<%= hfCopyAttachment.ClientID %>');
                            if (hfAttach) {
                                hfAttach.value = (answer === 'yes') ? '1' : '0';
                            }
                            copyStage = 'mdr';
                            var question = document.getElementById('divCopyQuestion');
                            if (question) {
                                question.innerText = '是否複製憑證明細？';
                            }
                            return;
                        }
                        if (copyStage === 'mdr') {
                            var hfMdr = document.getElementById('<%= hfCopyMDR.ClientID %>');
                            if (hfMdr) {
                                hfMdr.value = (answer === 'yes') ? '1' : '0';
                            }
                            hideCopyDialog();
                            __doPostBack('<%= btnCopyConfirm.UniqueID %>', '');
                        }
                    }

                    function hideCopyDialog() {
                        var modal = document.getElementById('divCopyModal');
                        if (modal) {
                            modal.style.display = 'none';
                        }
                    }
                </script>

                <asp:GridView ID="gvResults" runat="server" AutoGenerateColumns="False" CssClass="gridview"
                    AllowPaging="True" DataKeyNames="jID,CreateBy"
                    OnPageIndexChanging="gvResults_PageIndexChanging"
                    OnRowDataBound="gvResults_RowDataBound"
                    OnRowCommand="gvResults_RowCommand">
                    <Columns>
                        <asp:TemplateField HeaderText="動作">
                            <ItemTemplate>
                                <asp:LinkButton ID="lbtnCopy" runat="server" Text="複製"
                                    CssClass="btn btn-secondary btn-grid"
                                    CommandName="CopyDoc"
                                    CommandArgument='<%# Container.DataItemIndex %>'
                                    OnClientClick='<%# "return showCopyDialog(" & Container.DataItemIndex & ");" %>'>
                                </asp:LinkButton>
                            </ItemTemplate>
                            <ItemStyle HorizontalAlign="Center" Width="70px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="jID">
                            <ItemTemplate>
                                <asp:HyperLink ID="hlJID" runat="server" CssClass="link-jid"
                                    NavigateUrl='<%# GetDocumentUrl(Eval("jID")) %>'
                                    Text='<%# Eval("jID") %>'></asp:HyperLink>
                            </ItemTemplate>
                            <ItemStyle HorizontalAlign="Center" Width="60px" />
                        </asp:TemplateField>
                        <asp:BoundField DataField="CardName" HeaderText="供應商名稱" />
                        <asp:BoundField DataField="InvNum" HeaderText="AP單號" />
                        <asp:BoundField DataField="U_PID" HeaderText="簽核系統PID" />
                        <asp:TemplateField HeaderText="文件狀態">
                            <ItemTemplate>
                                <span class='<%# "badge status-" & Eval("ApprovalStatus") %>'>
                                    <%# GetStatusText(Eval("ApprovalStatus").ToString()) %>
                                </span>
                            </ItemTemplate>
                            <ItemStyle HorizontalAlign="Center" Width="80px" />
                        </asp:TemplateField>
                        <asp:BoundField DataField="DocDate" HeaderText="文件日期" DataFormatString="{0:yyyy-MM-dd}" />
                        <asp:TemplateField HeaderText="放行狀態">
                            <ItemTemplate>
                                <%# If(Eval("IsApproved") IsNot Nothing AndAlso Eval("IsApproved").ToString() = "Y", "已放行", "未放行") %>
                            </ItemTemplate>
                            <ItemStyle HorizontalAlign="Center" Width="70px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="備註">
                            <ItemTemplate>
                                <div class="remarks-cell" title='<%# Eval("Comments") %>'>
                                    <%# TruncateRemarks(Eval("Comments").ToString(), 20) %>
                                </div>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="CreateBy" HeaderText="建立者" />
                    </Columns>
                    <PagerStyle CssClass="pager" />
                    <EmptyDataTemplate>
                        <div class="empty-data">
                            請輸入篩選條件後按「查詢」
                        </div>
                    </EmptyDataTemplate>
                </asp:GridView>

            </ContentTemplate>
        </asp:UpdatePanel>
    </div>

    <!-- Hidden controls for modal -->
    <asp:Button ID="btnDummyRemarks" runat="server" style="display:none" />
    <ajaxToolkit:ModalPopupExtender ID="mpeRemarks" runat="server" TargetControlID="btnDummyRemarks"
        PopupControlID="pnlRemarks" BackgroundCssClass="modal-overlay" CancelControlID="btnCloseRemarks" />
    <asp:Panel ID="pnlRemarks" runat="server" CssClass="modal-content" style="display:none;">
        <div class="modal-footer">
            <asp:Button ID="btnCloseRemarks" runat="server" Text="關閉" CssClass="btn btn-secondary" />
        </div>
    </asp:Panel>
</asp:Content>
