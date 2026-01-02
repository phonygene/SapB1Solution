<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/MySite1.Master" CodeBehind="Home.aspx.vb" Inherits="MgmSP.Home" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <%--<link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,300;0,400;1,300;1,400&family=DM+Sans:wght@400;500&display=swap" rel="stylesheet" />--%>
    <link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,300;0,400;0,600;0,700;1,300;1,400;1,600;1,700&family=DM+Sans:wght@400;500&display=swap" rel="stylesheet" />

    <style>
        .content-area {
            /*background: transparent;*/
            background: linear-gradient(160deg, #171c30 0%, #1e2440 30%, #252d48 50%, #1e2440 70%, #171c30 100%);
            display: flex;
            justify-content: center;
            align-items: center;
        }
        .welcome-panel {
            text-align: center;
            opacity: 1 !important;
          filter: none !important;
          text-shadow: none !important;
          mix-blend-mode: normal !important;
        }
        .welcome-title {
            font-family: "Cormorant Garamond", Georgia, serif;
            font-size: 26px;
            font-weight: 600;
            font-style: italic;
            letter-spacing: 0.35em;
            /*color: rgba(180, 190, 215, 0.6);*/
            /*color: #1a1f35;*/
            color: #f0f8ff;
            text-transform: uppercase;
            margin-bottom: 16px;
            opacity: 1 !important;
          filter: none !important;
          text-shadow: none !important;
          mix-blend-mode: normal !important;
        }
        .welcome-user {
            font-family: "Cormorant Garamond", Georgia, serif;
            font-size: 36px;
            font-weight: 700;
            /*color: rgba(235, 240, 250, 0.85);*/
            /*color: #1a1f35;*/
            color: #f0f8ff;
            letter-spacing: 0.08em;
            opacity: 1 !important;
          filter: none !important;
          text-shadow: none !important;
          mix-blend-mode: normal !important;
        }

/*        .welcome-panel,
        .welcome-title,
        .welcome-user,
        .welcome-user span {
          opacity: 1 !important;
          filter: none !important;
          text-shadow: none !important;
          mix-blend-mode: normal !important;
        }*/
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div class="welcome-panel">
        <div class="welcome-title">Welcome</div>
        <div class="welcome-user"><asp:Label ID="lblUserName" runat="server" Text=""></asp:Label></div>
    </div>
</asp:Content>
