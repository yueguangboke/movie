<%
'*********************************************************
' 作者：未寒
' 邮箱：im@imzhou.com 
'*********************************************************
%>
<%
Call RegisterPlugin("SimplePro","ActivePlugin_SimplePro")

Function ActivePlugin_SimplePro()
	'配置初始化
	Dim c
	Set c = New TConfig
	c.Load("SimplePro")
	If c.Exists("vesion")=False Then
		c.Write "vesion","2.1"
		c.Write "Color","#0099CC"
		c.Save
	End If 

    Call Add_Response_Plugin("Response_Plugin_Admin_Top",MakeTopMenu(1,"SimplePro主题配置",BlogHost & "zb_users/theme/SimplePro/plugin/editor.asp","aSimpleProManage",""))
End Function

Function InstallPlugin_SimplePro()
	Call SetBlogHint(Empty,Empty,True)
End Function
%>