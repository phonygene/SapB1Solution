<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="FeatureMaintenance.aspx.vb" Inherits="MgmSP.FeatureMaintenance" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">

<head runat="server">
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>功能維護中</title>
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

        .feature-name {
            font-size: 18px;
            color: #f0c040;
            margin-bottom: 15px;
            font-weight: 500;
        }

        .maintenance-message {
            font-size: 15px;
            line-height: 1.8;
            color: rgba(180, 190, 215, 0.8);
            white-space: pre-line;
            margin-bottom: 30px;
        }

        .back-btn {
            display: inline-block;
            padding: 14px 32px;
            font-size: 14px;
            font-weight: 500;
            font-family: "DM Sans", sans-serif;
            letter-spacing: 0.06em;
            background: rgba(230, 235, 250, 0.95);
            border: none;
            border-radius: 8px;
            color: #1a1f35;
            text-decoration: none;
            cursor: pointer;
            transition: background 0.2s ease, transform 0.1s ease;
        }

        .back-btn:hover {
            background: #ffffff;
        }

        .back-btn:active {
            transform: scale(0.98);
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
                <img src="images/maintenance.svg" alt="維護中" class="maintenance-image"
                     onerror="this.style.display='none'" />

                <h1 class="maintenance-title">功能維護中</h1>
                
                <p class="feature-name">
                    <asp:Literal ID="litFeatureName" runat="server"></asp:Literal>
                </p>

                <p class="maintenance-message">
                    <asp:Literal ID="litMaintenanceNote" runat="server"></asp:Literal>
                </p>
                
                <a href="Home.aspx" class="back-btn">返回首頁</a>
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
