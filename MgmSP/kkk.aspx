<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="kkk.aspx.vb" Inherits="MgmSP.kkk" %>

<!DOCTYPE html>

<html lang="zh"> 
	<head>
	    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css"> 
	</head> 
	<body> 
	 
	<form id="form1" runat="server" autocomplete="off"> 
	    資料庫時間 : <asp:Label id="lb_db_time" runat="server" /><br> 
	    系統時間 : <asp:Label id="lb_sys_time" runat="server" /><br> 
	    結果 : <asp:Label id="lb_memo" runat="server" /> 
	</form> 
	 
	<!-- Modal --> 
	<div id="myModal" class="modal fade" role="dialog"> 
	    <div class="modal-dialog"> 
	        <div class="modal-content"> 
	            <div class="modal-header"> 
	                <button type="button" class="close" data-dismiss="modal">&times;</button> 
	                <h4 class="modal-title">訊息 : </h4> 
	            </div> 
	 
	            <div class="modal-body"> 
	                <span class="text-info">這是自動提示訊息</span> 
	            </div> 
	 
	            <div class="modal-footer"> 
	                <button type="button" class="btn btn-default" data-dismiss="modal">確定 (Close)</button> 
	            </div> 
	        </div> 
	    </div> 
	</div> 
	 
	 
	 
	<script src="https://ajax.googleapis.com/ajax/libs/jquery/1/jquery.min.js"></script> 
	<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js" integrity="sha384-Tc5IQib027qvyjSMfHjOMaLkfuWVxZxUPnCJA7l2mCWNIpG9mGCD8wGNIcPD7Txa" crossorigin="anonymous"></script> 
	 
	<asp:Literal id="custom_script" runat="server" /> 
	 
	</body> 
	</html> 
	 
	<Script Language="VB" runat="server"> 


	</Script> 

