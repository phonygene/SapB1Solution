<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Index.aspx.vb" Inherits="MgmSP.WebForm1" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">

<head runat="server">
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>JET Enterprise Platform</title>
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
            max-width: 480px;
            padding: 0 20px;
            display: flex;
            flex-direction: column;
            align-items: center;
        }

        .logo-section {
            width: 100%;
            margin-bottom: 48px;
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

        .user-info {
            margin-bottom: 36px;
            text-align: center;
        }

        .user-greeting {
            font-size: 13px;
            color: rgba(180, 190, 215, 0.5);
            letter-spacing: 0.02em;
        }

        .user-name {
            font-size: 16px;
            color: rgba(235, 240, 250, 0.9);
            margin-top: 4px;
            font-weight: 500;
        }

        .menu-section {
            width: 100%;
        }

        .menu-btn {
            display: block;
            width: 100%;
            padding: 16px;
            margin-bottom: 12px;
            font-size: 14px;
            font-weight: 500;
            font-family: "DM Sans", sans-serif;
            letter-spacing: 0.04em;
            background: rgba(230, 235, 250, 0.95);
            border: none;
            border-radius: 8px;
            color: #1a1f35;
            cursor: pointer;
            text-decoration: none;
            text-align: center;
            transition: background 0.2s ease, transform 0.1s ease;
        }

        .menu-btn:hover {
            background: #ffffff;
        }

        .menu-btn:active {
            transform: scale(0.98);
        }

        .menu-btn-secondary {
            background: rgba(255, 255, 255, 0.08);
            border: 1px solid rgba(255, 255, 255, 0.15);
            color: rgba(235, 240, 250, 0.95);
        }

        .menu-btn-secondary:hover {
            background: rgba(255, 255, 255, 0.12);
            border-color: rgba(200, 210, 235, 0.4);
        }

        .divider {
            width: 100%;
            height: 1px;
            background: rgba(180, 190, 215, 0.15);
            margin: 24px 0;
        }

        .logout-btn {
            display: block;
            width: 100%;
            padding: 14px;
            font-size: 13px;
            font-weight: 400;
            font-family: "DM Sans", sans-serif;
            letter-spacing: 0.03em;
            background: transparent;
            border: 1px solid rgba(255, 255, 255, 0.15);
            border-radius: 8px;
            color: rgba(180, 190, 215, 0.7);
            cursor: pointer;
            text-decoration: none;
            text-align: center;
            transition: all 0.2s ease;
        }

        .logout-btn:hover {
            background: rgba(255, 255, 255, 0.06);
            border-color: rgba(255, 255, 255, 0.25);
            color: rgba(235, 240, 250, 0.9);
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

        @media (max-width: 480px) {
            body {
                padding: 40px 20px;
            }

            .logo-img {
                max-width: 220px;
            }

            .menu-btn {
                padding: 14px;
                font-size: 13px;
            }
        }
    </style>
</head>

<body>
    <form id="form1" runat="server">
        <div class="main-content">
            <div class="logo-section">
                <img src="usermgm/images/logo.svg" alt="Logo" class="logo-img" />
                <span class="tagline">Precision With Motion</span>
            </div>

            <div class="user-info">
                <div class="user-greeting">Welcome</div>
                <div class="user-name">
                    <asp:Label ID="lblUserName" runat="server" Text=""></asp:Label>
                </div>
            </div>

            <div class="menu-section">
                <a href="ExpenseClaimList.aspx" class="menu-btn">Expense Claim</a>
                <a href="signoff/signofftodo.aspx" class="menu-btn">Pending Approval</a>
                <a href="DocumentSearch.aspx" class="menu-btn menu-btn-secondary">Document Search</a>

                <div class="divider"></div>

                <a href="usermgm/logout.aspx" class="logout-btn">Sign Out</a>
            </div>

            <div class="footer">
                <div class="footer-line"></div>
                <span class="footer-text">Enterprise Platform</span>
                <div class="footer-line"></div>
            </div>
        </div>
    </form>
</body>

</html>
