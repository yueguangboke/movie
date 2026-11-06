<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    catalog.asp
'// 开始时间:    2005.02.11
'// 最后修改:    
'// 备    注:    目录
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #include file="../zb_users/c_option.asp" -->
<!-- #include file="../zb_system/function/c_function.asp" -->
<!-- #include file="../zb_system/function/c_system_lib.asp" -->
<!-- #include file="../zb_system/function/c_system_base.asp" -->
<!-- #include file="../zb_system/function/c_system_plugin.asp" -->
<!-- #include file="../zb_system/function/c_system_event.asp" -->
<!-- #include file="../zb_users/plugin/p_config.asp" -->
<%
Dim html

Call System_Initialize()

'plugin node
For Each sAction_Plugin_Catalog_Begin in Action_Plugin_Catalog_Begin
	If Not IsEmpty(sAction_Plugin_Catalog_Begin) Then Call Execute(sAction_Plugin_Catalog_Begin)
Next

Dim objFSO
Dim objFile
Dim strFileName
Dim strFileTime
Dim isBuildFile

if Request.QueryString("tags")<>"" then
	if Request.QueryString("page")<>""then
		strFileName = "tag/" + Request.QueryString("tags") + "_" + Request.QueryString("page") +".html"
	else
		strFileName = "tag/" + Request.QueryString("tags") + ".html"
	end if
elseif Request.QueryString("date")<>"" then
	if Request.QueryString("page")<>""then
		strFileName = "date/" + Request.QueryString("date") + "-" + Request.QueryString("page") +".html"
	else
		strFileName = "date/" + Request.QueryString("date") + ".html"
	end if
elseif Request.QueryString("page")<>"" then
	strFileName = "page/index" + "_" + Request.QueryString("page") +".html"
else
	strFileName = ""
end If

Response.Status="301 Moved Permanently" 
Response.AddHeader "Location",BlogHost + strFileName
Response.end  




'plugin node
For Each sAction_Plugin_Catalog_End in Action_Plugin_Catalog_End
	If Not IsEmpty(sAction_Plugin_Catalog_End) Then Call Execute(sAction_Plugin_Catalog_End)
Next

Call System_Terminate()


If Err.Number<>0 then
	Call ShowError(0)
End If
%>
<!--<%=RunTime()%>ms-->