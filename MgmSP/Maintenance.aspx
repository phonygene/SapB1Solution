<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Maintenance.aspx.vb" Inherits="MgmSP.Maintenance" %>

    <!DOCTYPE html>
    <html xmlns="http://www.w3.org/1999/xhtml">

    <head runat="server">
        <meta charset="UTF-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title>系統維護中</title>
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
                padding: 60px 20px;
            }

            .main-content {
                width: 100%;
                max-width: 500px;
                padding: 0 40px;
                display: flex;
                flex-direction: column;
                align-items: center;
                text-align: center;
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

            .maintenance-container {
                width: 100%;
                padding: 40px 30px;
                background: rgba(255, 255, 255, 0.05);
                border: 1px solid rgba(255, 255, 255, 0.1);
                border-radius: 16px;
            }

            /* ========== 維護圖片 ========== */
            /* 要替換圖片，只需修改 maintenance-image 的 src 即可 */
            .maintenance-image {
                max-width: 200px;
                max-height: 200px;
                margin-bottom: 30px;
                opacity: 0.9;
            }

            .maintenance-title {
                font-size: 24px;
                font-weight: 500;
                color: rgba(235, 240, 250, 0.95);
                margin-bottom: 20px;
                letter-spacing: 0.05em;
            }

            .maintenance-message {
                font-size: 15px;
                line-height: 1.8;
                color: rgba(180, 190, 215, 0.8);
                white-space: pre-line;
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

            .admin-login-section {
                margin-top: 30px;
                padding-top: 20px;
                border-top: 1px solid rgba(255, 255, 255, 0.1);
            }

            .admin-login-btn {
                display: inline-block;
                padding: 12px 24px;
                font-size: 13px;
                font-weight: 500;
                font-family: "DM Sans", sans-serif;
                letter-spacing: 0.04em;
                background: transparent;
                border: 1px solid rgba(240, 192, 64, 0.5);
                border-radius: 6px;
                color: rgba(240, 192, 64, 0.8);
                text-decoration: none;
                cursor: pointer;
                transition: all 0.2s ease;
            }

            .admin-login-btn:hover {
                background: rgba(240, 192, 64, 0.1);
                border-color: rgba(240, 192, 64, 0.8);
                color: rgba(240, 192, 64, 1);
            }
        </style>
    </head>

    <body>
        <form id="form1" runat="server">
            <div class="main-content">
                <!-- Logo Section -->
                <div class="logo-section">
                    <img src="usermgm/images/logo.svg" alt="Logo" class="logo-img" />
                    <span class="tagline">Precision With Motion</span>
                </div>

                <!-- Maintenance Content -->
                <div class="maintenance-container">
                    <!-- ========== 維護顯示圖 ========== -->
                    <!-- 要替換圖片，修改下方 src 路徑即可，例如: images/maintenance.png -->
                    <img src="images/maintenance.svg" alt="維護中" class="maintenance-image"
                        onerror="this.style.display='none'" />

                    <h1 class="maintenance-title">系統維護中</h1>

                    <!-- 維護文字描述 (從資料庫 OADM.MNote 讀取) -->
                    <p class="maintenance-message">
                        <asp:Literal ID="litMaintenanceNote" runat="server"></asp:Literal>
                    </p>

                    <!-- 系統管理員登入區塊 -->
                    <div class="admin-login-section">
                        <a href="AdminLogin.aspx" class="admin-login-btn">系統管理員登入 (開發者測試用)</a>
                    </div>
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