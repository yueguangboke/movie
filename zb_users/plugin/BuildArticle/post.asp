<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="..\..\c_option.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_function.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_lib.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_base.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_event.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_manage.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_plugin.asp" -->
<!-- #include file="..\p_config.asp" -->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("BuildArticle")=False Then Call ShowError(48)

Dim Req
Set Req=Request
Dim iPage,iMin,iMax,iAuth,iCate,bStat
iPage=Req("page")
iMin=Req("min")
iMax=Req("max")
iAuth=Req("auth_id")
iCate=Req("cate_id")
bStat=Req("stat")

Call CheckParameter(iPage,"int",1)
Call CheckParameter(iMin,"int",1)
Call CheckParameter(iMax,"int",1)
Call CheckParameter(bStat,"bool",True)

Select Case Req("type")
	Case 2
		
		Call BuildByMinMax(iPage,iMin,iMax,"",Array("a","a"))
		
	Case 3
	
		Call CheckParameter(iCate,"int",0)
		Call BuildByMinMax(iPage,iMin,iMax,"AND log_CateID="&iCate,Array("cate_id",iCate))
		
	Case 4
	
		Call CheckParameter(iAuth,"int",0)
		Call BuildByMinMax(iPage,iMin,iMax,"AND log_AuthorID="&iAuth,Array("auth_id",iAuth))
		
	Case 5
	
		StarTime=Timer()
		Call MakeBlogRebuild_Core
		ExitError Array(Array("time",RunTime()),Array("msg","首页和侧栏重建完成，用时"&RunTime&"ms"),Array("ok",True))
			
End Select


'If ZC_POST_STATIC_MODE<>"STATIC" Then
'	ExitError Array(Array("msg","不是静态模式"),Array("ok",true))
'End If

%>

<%
Sub BuildByMinMax(page,min,max,where,otherresult)

	StarTime=Timer()
	Dim aryData,strSQL,objRs,i,aryList(),strResult
	
	i=0
	
	Dim objArticle

	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.CursorType = adOpenKeyset
	objRS.LockType = adLockReadOnly
	objRS.ActiveConnection=objConn
	If Not ZC_MSSQL_ENABLE Then
	
		objRS.Source="SELECT [log_ID] FROM [blog_Article] WHERE [log_Level]>1 AND (log_ID BETWEEN "&min&" AND "&max&") " &where
		
	Else
	
		objRS.Source="SELECT [log_ID],[log_Tag],[log_CateID],[log_Title],[log_Intro],[log_Content],[log_Level],[log_AuthorID],[log_PostTime],[log_CommNums],[log_ViewNums],[log_TrackBackNums],[log_Url],[log_Istop],[log_Template],[log_FullUrl],[log_Type],[log_Meta] FROM [blog_Article] WHERE log_Level>1 AND (log_ID BETWEEN "&min&" AND "&max&") " &where 
		
	End If
	objRS.Open()

	If (Not objRS.bof) And (Not objRS.eof) Then

		objRS.PageSize = ZC_REBUILD_FILE_COUNT

		objRS.AbsolutePage = page

		Redim aryList(objRS.PageSize+1)
		
		For i = 1 To objRS.PageSize

			aryList(i)=objRs("log_ID")
			
			If Not ZC_MSSQL_ENABLE Then
			
				Call BuildArticle_BuildArticle(objRS("log_ID"),bStat)
				
			Else
			
				Call BuildArticle_BuildArticle_HighSpeed(Array(objRS(0),objRS(1),objRS(2),objRS(3),objRS(4),objRS(5),objRS(6),objRS(7),objRS(8),objRS(9),objRS(10),objRS(11),objRS(12),objRS(13),objRS(14),objRS(15),objRS(16),objRS(17)),bStat)
				
			End If
			
			objRS.MoveNext
			
			If objRS.EOF Then Exit For
			
			If Err.Number>0 Then ExitError Array(Array("ok",True),Array("msg","出现错误"&Err.Number&"："&Err.Description))
			
		Next
		
		'strResult=Join(aryList,",")
		
		'Do Until InStr(strResult,",,")<=0
		'	strResult=Replace(strResult,",,","")
		'Loop
	
		strResult=aryList(1) & "-" & aryList(i-1)
	
		ExitError Array(Array("page",page+1),Array("min",min),Array("max",max),Array("time",RunTime()),Array("msg","文章"&strResult&"重建完成，用时"&RunTime&"ms"),Array("dbmax",aryList(i-1)),Array("type",Req("type")),Array("stat",bStat),Array("ok",Not i>=ZC_REBUILD_FILE_COUNT),otherresult)

	Else
		
		ExitError Array(Array("msg","完成"),Array("ok",true))
		
	End If
	
	Set objRs=Nothing

	'aryData=BuildArticle_GetDoubleTop(page,ZC_REBUILD_FILE_COUNT,objConn.Execute("SELECT MAX(log_ID) FROM blog_Article WHERE 1=1 "&where)(0),"log_ID","blog_Article",","ORDER BY log_ID ASC","log_ID")
	'If aryData(0)=False Then ExitError Array(Array("msg","完成"),Array("ok",true))
	'
	'strSQL=aryData(1)
	'Response.Write strSQL&"<br/>"
	'Set objRs=objConn.Execute(strSQL)
	'Do Until objRs.EOF
	
		
	'	aryList(i)=objRs("log_ID")
	'	i=i+1
	'	objRs.MoveNext
		
	'Loop
	
	
	

	
End Sub


Sub ExitError(ary)
	Dim i,j
	For i=0 To Ubound(ary)
		j=j&""""&ary(i)(0)&""":"&IIf(VarType(ary(i)(1))=VarType(True),LCase(ary(i)(1)),""""&ary(i)(1)&"""")
		If i<Ubound(ary) Then j=j&","
	Next
	Response.Write "{"&j&"}":Response.End
End Sub
%>