<!-- #include file="function.asp" -->
<%
'注册插件
Call RegisterPlugin("BaiduRWUD","ActivePlugin_BaiduRWUD")
'挂口部分
Dim BaiduRWUDEnable
Function ActivePlugin_BaiduRWUD()
	Call Add_Action_Plugin("Action_Plugin_Edit_Article_Begin","Call BaiduRWUD_addForm()")
	Call Add_Action_Plugin("Action_Plugin_ArticlePst_Begin","Call BaiduRWUD_Main()")
End Function

Dim BaiduRWUD_cfg,BaiduRWUD_log
BaiduRWUD_cfg = BlogPath & "zb_users/PLUGIN/BaiduRWUD/"&ZC_BLOG_CLSID&".txt"
BaiduRWUD_log = BlogPath&"zb_users/PLUGIN/BaiduRWUD/log.txt"

Function BaiduRWUD_addForm()
  Call Add_Response_Plugin("Response_Plugin_Edit_Form3","<p><label>推送到百度:<input type=""checkbox"" name=""BaiduRWUDEnable"" id=""BaiduRWUDEnable"" value=""True"" checked/></label></p>")
End Function

Function BaiduRWUD_Main()
  BaiduRWUDEnable = IIF(Request.Form("BaiduRWUDEnable")="True",True,False)
	If BaiduRWUDEnable And PublicObjFSO.FileExists(BaiduRWUD_cfg) Then
		Call Add_Filter_Plugin("Filter_Plugin_PostArticle_Succeed","BaiduRWUD_Ping")
	End If
End Function

'发送数据
Function BaiduRWUD_Ping(ByRef objArticle)
  Dim xmlhttp,artUrl
  Set xmlhttp=Server.CreateObject("MSXML2.ServerXMLHTTP")
  xmlhttp.Open "POST",LoadFromFile(BaiduRWUD_cfg,"utf-8"),false
  xmlhttp.setRequestHeader "Content-Type", "text/plain"
  artUrl = objArticle.Url
  xmlhttp.send(artUrl)
  BaiduRWUD_Ping = xmlhttp.responseText
  Call SaveToFile(BaiduRWUD_log,BaiduRWUD_Ping,"UTF-8",False)
  Call SetBlogHint_Custom(BaiduRWUD_Ping&artUrl)
  Set xmlhttp=Nothing
End Function

Function InstallPlugin_BaiduRWUD()

	'用户激活插件之后的操作

End Function


Function UnInstallPlugin_BaiduRWUD()

	'用户停用插件之后的操作

End Function
%>