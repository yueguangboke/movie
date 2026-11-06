<%
'///////////////////////////////////////////////////////////////////////////////
'//插件名：FirstIMG
'//原名称：xiaoyi_FirstIMG
'//原作者：小一
'//作者：chemdisk
'//网站：www.chemdisk.com
'//提取文章中第一张图片，更新后支持Z-Blog 2.0版。
'//
'///////////////////////////////////////////////////////////////////////////////


'注册插件
Call RegisterPlugin("FirstIMG","ActivePlugin_FirstIMG")

Function ActivePlugin_FirstIMG()

	Call Add_Filter_Plugin("Filter_Plugin_TArticle_Export_TemplateTags","FirstIMG_func")

End Function


Function FirstIMG_func(ByRef aryTemplateTagsName,ByRef aryTemplateTagsValue)

		Dim c:c=UBOUND(aryTemplateTagsName)+1

		ReDim Preserve aryTemplateTagsName(c)
		ReDim Preserve aryTemplateTagsValue(c)
		
		Dim ArticleFirstIMG:ArticleFirstIMG=aryTemplateTagsValue(5)
		
		aryTemplateTagsName(c)="article/firstimg"
        Dim xiaoyiimgtemp:xiaoyiimgtemp = FirstImg_GetImgSrc(ArticleFirstIMG)
		If xiaoyiimgtemp <> "" Then
			aryTemplateTagsValue(c)="<img class=""firstimg"" src=""" & FirstImg_GetImgSrc(ArticleFirstIMG) & """/>"
		Else
		    aryTemplateTagsValue(c)="<img  class=""firstimg"" src=""<#ZC_BLOG_HOST#>zb_users/PLUGIN/FirstIMG/NoImages.jpg""/>"
		End IF
End Function


Function FirstImg_GetImgSrc(str)
Dim tmp, objRegExp, Matches, Match
Set objRegExp = New Regexp
objRegExp.IgnoreCase = True
objRegExp.Global = false
objRegExp.Pattern = "<img [^>]*?\bsrc=(['""]?)([^'""\x20]+)\1 [^>]*?>"
Set Matches = objRegExp.Execute(str)
For Each Match In Matches
tmp = tmp & Match.SubMatches(1)
Next
FirstImg_GetImgSrc = tmp
End Function 

%>