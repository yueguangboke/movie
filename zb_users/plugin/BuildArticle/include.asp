<!-- #include file="function.asp" -->
<%
'注册插件
Call RegisterPlugin("BuildArticle","ActivePlugin_BuildArticle")
'挂口部分
Function ActivePlugin_BuildArticle()

	Call Add_Response_Plugin("Response_Plugin_AskFileReBuild_SubMenu","<script>$(document).ready(function(){$('#edit').submit(function(){location.href=bloghost+'zb_users/plugin/buildarticle/main.asp?build=all';return false;})})</script>"&MakeSubMenu("分段重建",BlogHost&"zb_users/plugin/buildarticle/main.asp","m-left",False))
	
End Function


Function InstallPlugin_BuildArticle()

	'用户激活插件之后的操作
	
End Function


Function UnInstallPlugin_BuildArticle()

	'用户停用插件之后的操作
	
End Function
%>