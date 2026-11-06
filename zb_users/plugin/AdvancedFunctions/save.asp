<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("AdvancedFunctions")=False Then Call ShowError(48)
BlogTitle="高级侧栏"
Dim subCate
%>
<%
Set AdvancedFunctions_=AdvancedFunctions_Init()

Dim w

Dim attr2
Select Case Request.QueryString("act")
	Case "save"
		For Each attr2 In Request.Form
			w=Request.Form(attr2)
			If Left(attr2,2)="m_" Then
				Call CheckParameter(w,"int",0)
				If w<0 Then w=-w
				AdvancedFunctions_.config.orig.Write Right(attr2,Len(attr2)-2),w
			ElseIf Left(attr2,2)="s_" Then
				Call CheckParameter(w,"bool",True)
				AdvancedFunctions_.config.orig.Write Right(attr2,Len(attr2)-2),w
			End If	
			
			'Response.Write attr2
		Next
End Select
AdvancedFunctions_.config.orig.Save
SetBlogHint True,Empty,Empty
Response.Redirect "main.asp"
%>
<script language="javascript" runat="server">
var o=0;

</script>

