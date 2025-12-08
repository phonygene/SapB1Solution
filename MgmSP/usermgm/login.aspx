<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="login.aspx.vb" Inherits="MgmSP.login" %>

    <!DOCTYPE html>
    <html xmlns="http://www.w3.org/1999/xhtml">

    <head runat="server">
        <meta charset="UTF-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title>JET 登錄</title>
        <link
            href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,300;0,400;1,300;1,400&family=DM+Sans:wght@400;500&display=swap"
            rel="stylesheet" />
        <style>
            * {
                margin: 0;
                padding: 0;
                box-sizing: border-box;
            }

            html,
            body {
                min-height: 100vh;
            }

            body {
                background: linear-gradient(160deg, #171c30 0%, #1e2440 30%, #252d48 50%, #1e2440 70%, #171c30 100%);
                display: flex;
                flex-direction: column;
                align-items: center;
                justify-content: center;
                font-family: "DM Sans", -apple-system, sans-serif;
                padding: 60px 0;
            }

            .main-content {
                width: 100%;
                max-width: 420px;
                padding: 0 40px;
                display: flex;
                flex-direction: column;
                align-items: center;
            }

            .logo-section {
                width: 100%;
                margin-bottom: 40px;
                display: flex;
                flex-direction: column;
                align-items: center;
            }

            .logo-img {
                max-width: 280px;
                max-height: 100px;
                object-fit: contain;
            }

            .tagline {
                margin-top: 16px;
                font-family: "Cormorant Garamond", Georgia, serif;
                font-size: 14px;
                font-weight: 300;
                font-style: italic;
                letter-spacing: 0.35em;
                color: rgba(180, 190, 215, 0.6);
                text-transform: uppercase;
            }

            .login-form {
                width: 100%;
            }

            .input-group {
                margin-bottom: 18px;
            }

            .form-input {
                width: 100%;
                padding: 16px 20px;
                font-size: 15px;
                font-family: "DM Sans", sans-serif;
                background: rgba(255, 255, 255, 0.08);
                border: 1px solid rgba(255, 255, 255, 0.15);
                border-radius: 8px;
                color: rgba(235, 240, 250, 0.95);
                outline: none;
                transition: border-color 0.2s ease, background 0.2s ease;
            }

            .form-input::placeholder {
                color: rgba(180, 190, 215, 0.5);
            }

            .form-input:focus {
                background: rgba(255, 255, 255, 0.12);
                border-color: rgba(200, 210, 235, 0.4);
            }

            .form-select {
                width: 100%;
                padding: 16px 20px;
                font-size: 15px;
                font-family: "DM Sans", sans-serif;
                background: rgba(255, 255, 255, 0.08);
                border: 1px solid rgba(255, 255, 255, 0.15);
                border-radius: 8px;
                color: rgba(235, 240, 250, 0.95);
                outline: none;
                cursor: pointer;
                appearance: none;
                background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'%3E%3Cpath fill='rgba(180,190,215,0.6)' d='M6 8L1 3h10z'/%3E%3C/svg%3E");
                background-repeat: no-repeat;
                background-position: right 16px center;
                padding-right: 44px;
                transition: border-color 0.2s ease, background 0.2s ease;
            }

            .form-select:focus {
                background-color: rgba(255, 255, 255, 0.12);
                border-color: rgba(200, 210, 235, 0.4);
            }

            .form-select option {
                background: #1a1f35;
                color: rgba(235, 240, 250, 0.95);
            }

            .form-label {
                display: block;
                font-size: 13px;
                color: rgba(180, 190, 215, 0.7);
                margin-bottom: 8px;
                letter-spacing: 0.02em;
            }

            .login-btn {
                width: 100%;
                padding: 16px;
                margin-top: 8px;
                font-size: 14px;
                font-weight: 500;
                font-family: "DM Sans", sans-serif;
                letter-spacing: 0.06em;
                background: rgba(230, 235, 250, 0.95);
                border: none;
                border-radius: 8px;
                color: #1a1f35;
                cursor: pointer;
                transition: background 0.2s ease, transform 0.1s ease;
            }

            .login-btn:hover {
                background: #ffffff;
            }

            .login-btn:active {
                transform: scale(0.98);
            }

            .error-msg {
                color: #ff6b6b;
                font-size: 13px;
                margin-bottom: 12px;
                text-align: center;
                min-height: 18px;
            }

            .footer {
                margin-top: 48px;
                display: flex;
                justify-content: center;
                align-items: center;
                gap: 16px;
            }

            .footer-line {
                width: 50px;
                height: 1px;
                background: rgba(180, 190, 215, 0.2);
            }

            .footer-text {
                font-size: 10px;
                color: rgba(180, 190, 215, 0.4);
                letter-spacing: 0.2em;
                text-transform: uppercase;
            }

            .hidden {
                display: none;
            }
        </style>
        <script type="text/javascript">
            // 防止登入按鈕連續點擊
            var isSubmitting = false;
            function preventDoubleClick(btn) {
                if (isSubmitting) {
                    return false;
                }
                isSubmitting = true;
                setTimeout(function () {
                    btn.disabled = true;
                    btn.value = '登錄中...';
                }, 10);
                return true;
            }
        </script>
    </head>

    <body>
        <form id="form1" runat="server">
            <div class="main-content">
                <!-- Logo Section -->
                <div class="logo-section">
                    <!-- ========== LOGO PATH ========== -->
                    <!-- Change the src path below to your logo image -->
                    <img src="images/logo.svg" alt="Logo" class="logo-img" />
                    <span class="tagline">Precision With Motion</span>
                </div>

                <!-- Login Form -->
                <div class="login-form">
                    <!-- Hidden Server Field -->
                    <div class="hidden">
                        <asp:Label ID="Label3" runat="server" Text="Server"></asp:Label>
                        <asp:TextBox ID="ServerText" runat="server" AutoPostBack="True">.\SQLEXPRESS2008R2</asp:TextBox>
                    </div>

                    <!-- Hidden Database Dropdown -->
                    <div class="hidden">
                        <asp:Label ID="Label4" runat="server" Text="資料庫"></asp:Label>
                        <asp:DropDownList ID="DDLServer" runat="server" AutoPostBack="True"></asp:DropDownList>
                    </div>

                    <!-- Warehouse/Model Selection - 隱藏，此版本不使用生產製造功能 -->
                    <div class="hidden">
                        <asp:DropDownList ID="DDLWhs" runat="server" CssClass="form-select"></asp:DropDownList>
                        <asp:Label ID="Label5" runat="server" Text="" Visible="false"></asp:Label>
                    </div>

                    <!-- Account Input -->
                    <div class="input-group">
                        <asp:TextBox ID="idtxt" runat="server" CssClass="form-input" placeholder="帳號"></asp:TextBox>
                    </div>

                    <!-- Password Input -->
                    <div class="input-group">
                        <asp:TextBox ID="pwdtxt" runat="server" TextMode="Password" CssClass="form-input"
                            placeholder="密碼"></asp:TextBox>
                    </div>

                    <!-- Error Message -->
                    <div class="error-msg">
                        <asp:Label ID="errmsg" runat="server" ForeColor="#FF6B6B"></asp:Label>
                    </div>

                    <!-- Login Button -->
                    <asp:Button ID="loginbtn" runat="server" Text="登錄" CssClass="login-btn"
                        OnClientClick="return preventDoubleClick(this);" />
                </div>

                <!-- Footer -->
                <div class="footer">
                    <div class="footer-line"></div>
                    <span class="footer-text">Enterprise Platform</span>
                    <div class="footer-line"></div>
                </div>
            </div>
        </form>
    </body>

    </html>