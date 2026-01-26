<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="AdminLogin.aspx.vb" Inherits="MgmSP.AdminLogin" %>

    <!DOCTYPE html>
    <html xmlns="http://www.w3.org/1999/xhtml">

    <head runat="server">
        <meta charset="UTF-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title>系統管理員登入</title>
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

            .admin-badge {
                background: linear-gradient(135deg, #f0c040, #e09800);
                color: #1a1f35;
                padding: 8px 20px;
                border-radius: 20px;
                font-size: 12px;
                font-weight: 600;
                letter-spacing: 0.1em;
                margin-bottom: 30px;
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

            .login-btn {
                width: 100%;
                padding: 16px;
                margin-top: 8px;
                font-size: 14px;
                font-weight: 500;
                font-family: "DM Sans", sans-serif;
                letter-spacing: 0.06em;
                background: linear-gradient(135deg, #f0c040, #e09800);
                border: none;
                border-radius: 8px;
                color: #1a1f35;
                cursor: pointer;
                transition: background 0.2s ease, transform 0.1s ease;
            }

            .login-btn:hover {
                background: linear-gradient(135deg, #ffd060, #f0a800);
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

            .back-link {
                display: block;
                margin-top: 20px;
                text-align: center;
                color: rgba(180, 190, 215, 0.6);
                font-size: 13px;
                text-decoration: none;
                transition: color 0.2s ease;
            }

            .back-link:hover {
                color: rgba(200, 210, 235, 0.9);
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
        </style>
        <script type="text/javascript">
            var isSubmitting = false;
            function preventDoubleClick(btn) {
                if (isSubmitting) {
                    return false;
                }
                isSubmitting = true;
                setTimeout(function () {
                    btn.disabled = true;
                    btn.value = '登入中...';
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
                    <img src="usermgm/images/logo.svg" alt="Logo" class="logo-img" />
                    <span class="tagline">Precision With Motion</span>
                </div>

                <div class="admin-badge">系統管理員專用</div>

                <!-- Login Form -->
                <div class="login-form">
                    <!-- Account Input -->
                    <div class="input-group">
                        <asp:TextBox ID="txtUserId" runat="server" CssClass="form-input" placeholder="管理員帳號">
                        </asp:TextBox>
                    </div>

                    <!-- Password Input -->
                    <div class="input-group">
                        <asp:TextBox ID="txtPassword" runat="server" TextMode="Password" CssClass="form-input"
                            placeholder="密碼"></asp:TextBox>
                    </div>

                    <!-- Error Message -->
                    <div class="error-msg">
                        <asp:Label ID="lblError" runat="server" ForeColor="#FF6B6B"></asp:Label>
                    </div>

                    <!-- Login Button -->
                    <asp:Button ID="btnLogin" runat="server" Text="管理員登入" CssClass="login-btn"
                        OnClientClick="return preventDoubleClick(this);" />

                    <!-- Back Link -->
                    <a href="Maintenance.aspx" class="back-link">← 返回維護頁面</a>
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