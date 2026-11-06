<%
'****************************************
' BuildArticle 子菜单
'****************************************
Function BuildArticle_SubMenu(id)
	Dim aryName,aryPath,aryFloat,aryInNewWindow,i
	aryName=Array("首页")
	aryPath=Array("main.asp")
	aryFloat=Array("m-left")
	aryInNewWindow=Array(False)
	For i=0 To Ubound(aryName)
		BuildArticle_SubMenu=BuildArticle_SubMenu & MakeSubMenu(aryName(i),aryPath(i),aryFloat(i)&IIf(i=id," m-now",""),aryInNewWindow(i))
	Next
End Function

Function BuildArticle_GetDoubleTop(intPage,PageSize,intMax,FieldName,TableName,whereData,order_by,id_field)
	Dim PageSize2,aryResult(1)
	PageSize2=PageSize
	If intPage<1 Then intPage=1
	If PageSize*intPage>intMax Then
		PageSize2=CLng(intMax Mod PageSize)
		If PageSize*(intPage-1)+PageSize2>intMax Then
			aryResult(0)=False
			aryResult(1)=False
			BuildArticle_GetDoubleTop=aryResult
			Exit Function
		End If
	End If
	aryResult(0)=PageSize2
	aryResult(1)="SELECT * FROM (SELECT TOP "&PageSize2&" *  FROM (SELECT TOP "&(PageSize*intPage)&" "&FieldName&" FROM ["&TableName&"]  "&IIf(whereData="","","WHERE ("&whereData&")")&" ORDER BY "&id_field&" ASC) As [Test] ORDER BY "&id_field&" DESC)"&" "&IIf(order_by="",""," As [test] "&order_by)&""
	BuildArticle_GetDoubleTop=aryResult
End Function

Function BuildArticle_BuildArticle(intID,bolStat)

	Dim objArticle
	Set objArticle=New TArticle

	If objArticle.LoadInfoByID(intID) Then
		If objArticle.Export(ZC_DISPLAY_MODE_ALL) Then
			objArticle.SaveCache
			If bolStat Then objArticle.Statistic
			objArticle.Build
			objArticle.Save
		End If

		BuildArticle_BuildArticle=True

	End If

	Set objArticle=Nothing

End Function

Function BuildArticle_BuildArticle_HighSpeed(ary,bolStat)

	Dim objArticle
	Set objArticle=New TArticle

	If objArticle.LoadInfoByArray(ary) Then
		If objArticle.Export(ZC_DISPLAY_MODE_ALL) Then
			objArticle.SaveCache
			If bolStat Then objArticle.Statistic
			objArticle.Build
			objArticle.Save
		End If

		BuildArticle_BuildArticle_HighSpeed=True

	End If

	Set objArticle=Nothing

End Function

%>