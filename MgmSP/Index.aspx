<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Index.aspx.vb" Inherits="MgmSP.WebForm1" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>JET Enterprise Platform</title>
    <link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,300;0,400;1,300;1,400&family=DM+Sans:wght@400;500&display=swap" rel="stylesheet" />
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        html, body {
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
            max-width: 420px;
            padding: 0 40px;
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
        .login-btn {
            width: 100%;
            padding: 16px;
            font-size: 14px;
            font-weight: 500;
            font-family: "DM Sans", sans-serif;
            letter-spacing: 0.06em;
            background: rgba(230, 235, 250, 0.95);
            border: none;
            border-radius: 8px;
            color: #1a1f35;
            cursor: pointer;
            text-decoration: none;
            text-align: center;
            display: block;
            transition: background 0.2s ease, transform 0.1s ease;
        }
        .login-btn:hover {
            background: #ffffff;
        }
        .login-btn:active {
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
            <div class="logo-section">
                <img src="usermgm/images/logo.svg" alt="Logo" class="logo-img" />
                <span class="tagline">Precision With Motion</span>
            </div>

            <a href="usermgm/login.aspx" class="login-btn">登入</a>

            <div class="footer">
                <div class="footer-line"></div>
                <span class="footer-text">Enterprise Platform</span>
                <div class="footer-line"></div>
            </div>
        </div>
    </form>
</body>
</html>
