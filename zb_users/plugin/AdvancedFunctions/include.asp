<!-- #include file="functions.asp"-->
<%
Const AdvancedFunctions_Empty=Empty
Dim AdvancedFunctions_
'*********************************************************
' 挂口: 注册插件和接口
'*********************************************************
'注册插件
Call RegisterPlugin("AdvancedFunctions","ActivePlugin_AdvancedFunctions")
'挂口部分


Function ActivePlugin_AdvancedFunctions()
	Call Add_Action_Plugin("Action_Plugin_MakeBlogReBuild_Core_Begin","Set AdvancedFunctions_=AdvancedFunctions_Init():AdvancedFunctions_.GoAll()")
	Call Add_Response_Plugin("Response_Plugin_FunctionMng_SubMenu",MakeSubMenu("高级侧栏","../../zb_users/plugin/advancedfunctions/main.asp","m-right",False))
End Function

Function InstallPlugin_AdvancedFunctions()
	objConn.Execute "DELETE FROM blog_Config WHERE conf_Name='AdvancedFunction'"
End Function

%>