<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/MySite1.Master" CodeBehind="Home.aspx.vb" Inherits="MgmSP.Home" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,300;0,400;1,300;1,400&family=DM+Sans:wght@400;500&display=swap" rel="stylesheet" />
    <style>
        .content-area {
            background: transparent;
            display: flex;
            justify-content: center;
            align-items: center;
        }
        .welcome-panel {
            text-align: center;
        }
        .welcome-title {
            font-family: "Cormorant Garamond", Georgia, serif;
            font-size: 18px;
            font-weight: 300;
            font-style: italic;
            letter-spacing: 0.35em;
            color: rgba(180, 190, 215, 0.6);
            text-transform: uppercase;
            margin-bottom: 16px;
        }
        .welcome-user {
            font-family: "Cormorant Garamond", Georgia, serif;
            font-size: 32px;
            font-weight: 400;
            color: rgba(235, 240, 250, 0.85);
            letter-spacing: 0.08em;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div class="welcome-panel">
        <div class="welcome-title">Welcome</div>
        <div class="welcome-user"><asp:Label ID="lblUserName" runat="server" Text=""></asp:Label></div>
    </div>
</asp:Content>
