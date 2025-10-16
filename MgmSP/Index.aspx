<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/MySite1.Master" CodeBehind="Index.aspx.vb" Inherits="MgmSP.WebForm1" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <!-- Ver: V2.4.8 2025/02/18 -->
    <!--若有新相機, 要在sap 此料號之項目主檔中的相機屬性勾選 -->
    <!--有新機型 要在sap 上自訂表格之UMMD 把資料填上 -->
    <!--若有新骨架料號 要在sap 上自訂表格之UMFM 把資料填上-->

    <!--在Local and 測試資料庫debug 後 , 欲發佈置網站時, 請做下列動作-->
    <!--CommUtil.vb 1. 把 Sub SendMail 中之 ToAddress = "ron@jettech.com.tw" mark起來disable-->
    <!--global.asax 之 Application("http") 改成適合的 -->
    <!--刪除 SapB1Solution\MgmSP\AttachFile 目錄下所有檔案或子目錄(這些是debug所留下來的) -->
    <!-- login.aspx.vb 要把 JTTST1 改為 JTSTD-->
    <!-- Web.config 要把JTTS1 改為 JTSTD 之連接-->
    <!-- MySite.Master 把Title "Jet工廠資訊系統 VX.x.x YYYY" X.x.x 改為發佈版本 , YYYY 改為Working-->
    <!-- 此項只是說明,不需修改 -- Application("http") 不能出現在單獨 .vb 之程式,否則會執行到那沒動作, 故在CommSignOff.vb 中有用到此 , 要以參數帶進-->

    <!-- 以下已由global.asax 設置Application("http") 代替-->
    <!-- OK 要debug 簽核郵件時 cLsignoff.aspx.vb 之 1. url 改掉-->
    <!-- OK molist.aspx.vb 之btnx_Click 中之url要改對 -->
    <!-- OK signofftodo.aspx.vb 之ShowSighOffPdf url 要改 -->
    <!-- OK iqc.aspx.vb 之url 要改 -->
    <!-- OK printform.aspx.vb 之url 要改 -->
    <!-- OK CommSignOff.vb 1.href = 改成適合的-->

    <!-- sessionState cookieless="true"-->



    <!-- jtdb 加入rd000 rd100-->
    <p>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </p>
    </asp:Content>
