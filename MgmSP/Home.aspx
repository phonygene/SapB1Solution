<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/MySite1.Master" CodeBehind="Home.aspx.vb" Inherits="MgmSP.Home" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style>
        /* 首頁保持深色背景風格 */
        .content-area {
            background: transparent;
            display: flex;
            justify-content: center;
            align-items: center;
        }
        .welcome-panel {
            text-align: center;
            padding: 60px 40px;
        }
        .welcome-title {
            font-family: "DM Sans", sans-serif;
            font-size: 32px;
            font-weight: 400;
            color: rgba(235, 240, 250, 0.9);
            margin-bottom: 12px;
            letter-spacing: 0.02em;
        }
        .welcome-user {
            font-family: "DM Sans", sans-serif;
            font-size: 15px;
            color: rgba(180, 190, 215, 0.6);
            letter-spacing: 0.01em;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div class="welcome-panel">
        <div class="welcome-title">Welcome</div>
        <div class="welcome-user">
            <asp:Label ID="lblUserName" runat="server" Text=""></asp:Label>
        </div>
    </div>
</asp:Content>
