<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/MySite1.Master" CodeBehind="Home.aspx.vb" Inherits="MgmSP.Home" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style>
        .welcome-panel {
            background: rgba(255, 255, 255, 0.05);
            border: 1px solid rgba(255, 255, 255, 0.1);
            border-radius: 8px;
            padding: 40px;
            text-align: center;
            max-width: 600px;
        }
        .welcome-title {
            font-size: 24px;
            font-weight: 500;
            color: rgba(235, 240, 250, 0.95);
            margin-bottom: 8px;
        }
        .welcome-user {
            font-size: 16px;
            color: rgba(180, 190, 215, 0.7);
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
