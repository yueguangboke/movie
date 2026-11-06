<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 插件制作:    ZSXSOFT
'///////////////////////////////////////////////////////////////////////////////
%>
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
If CheckPluginState("AdvancedConfig")=False Then Call ShowError(48)
BlogTitle="AdvancedConfig-高级配置中心"
Dim i
If Request.QueryString("act")="save" Then
	For i=1 To BlogConfig.Count
		If Not IsEmpty(Request.Form(BlogConfig.Meta.Names(i))) Then
			BlogConfig.Write BlogConfig.Meta.Names(i),Request.Form(BlogConfig.Meta.Names(i))
		End If
		
	Next
	BlogConfig.Save
	SaveConfig2Option
	Call SetBlogHint(True,Empty,Empty)
	Response.Redirect "main.asp"
	
End If

			
Call js_getAllConfigs
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<style type="text/css">
.beta td[_type="name"] {
	color: blue
}
.throwed {
	display: none
}
.hidden {
	display: none
}
tr {
	height: 32px
}
input[type="text"]{
	width: 98%;
	*width: 490px
}
textarea{
	width: 600px
}
</style>
<script type="text/javascript">
function bmx2table(){
	var class_=new Array("color2","color3","color4");
	var j=$("table tr:has(th)").addClass("color1");
    $("table").each(function(){
 		if(j.length==0){class_[1]="color2";class_[0]="color3";} 
		$(this).find("tr:not(:has(th)):not(.throwed):even").removeClass(class_[0]).addClass(class_[1]);
		$(this).find("tr:not(:has(th)):not(.throwed):odd").removeClass(class_[1]).addClass(class_[0]);
	})
	$("tr:not(:has(th))").mouseover(function(){$(this).addClass(class_[2])}).mouseout(function(){$(this).removeClass(class_[2])}); 
};
</script>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"><%=Response_Plugin_SettingMng_SubMenu%></div>
          <div id="divMain2"> 
            <script type="text/javascript">ActiveTopMenu("topmenu2");$(".SubMenu a[href*='zb_users/plugin/advancedconfig/main.asp']").find("span").addClass("m-now");</script>
            <form id="form1" name="form1" method="post" action="?act=save">
              <div class="content-box"><!-- Start Content Box -->
                
                <div class="content-box-header">
                  <ul class="content-box-tabs">
                    <li><a href="#tab_base" class="default-tab"><span>基础设置</span></a></li>
                    <li><a href="#tab_database"><span>数据库设置</span></a></li>
                    <li><a href="#tab_global"><span>全局设置</span></a></li>
                    <li><a href="#tab_page"><span>页面设置</span></a></li>
                    <li><a href="#tab_build"><span>静态化设置</span></a></li>
                    <li><a href="#tab_ubb"><span>UBB设置</span></a></li>
                    <li><a href="#tab_test"><span>测试项目</span></a></li>
                    <li><a href="#tab_others"><span>其他</span></a></li>
                    <li><a href="#tab_readonly"><span>只读</span></a></li>
                  </ul>
                  <div class="clear"></div>
                </div>
                <!-- End .content-box-header -->
                
                <div class="content-box-content">
                  <div class="tab-content default-tab" style='border:none;padding:0px;margin:0;' id="tab_base">
					<%=aryCfgList.base.join(vbCrlf)%>
                  </div>
                  
                  <div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab_database">
					<%=aryCfgList.database.join(vbCrlf)%>
                  </div>
                  
                  <div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab_global">
					<%=aryCfgList.global.join(vbCrlf)%>
                  </div>
                  
                  <div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab_page">
					<%=aryCfgList.page.join(vbCrlf)%>
                  </div>
                  
                  <div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab_build">
					<%=aryCfgList.build.join(vbCrlf)%>
                  </div>
                  
                  <div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab_ubb">
					<%=aryCfgList.ubb.join(vbCrlf)%>
                  </div>
                  
                  <div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab_test">
					<%=aryCfgList.test.join(vbCrlf)%>
                  </div>
                  
                  <div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab_others">
					<%=aryCfgList.others.join(vbCrlf)%>
                  </div>
                  
                  <div class="tab-content" style='border:none;padding:0px;margin:0;' id="tab_readonly">
					<%=aryCfgList.readonly.join(vbCrlf)%>
                  </div>
                </div>
              </div>

              <p>
                <input type="submit" value="提交" class="button"/>
              </p>
            </form>
          </div>
        </div>
        
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%Call System_Terminate()%>
<%

Function MakeTr(s,name,value,css)
	MakeTr="                                                      <tr"&IIf(Trim(css & " " & s.css)="",""," class="""&css & " " & s.css&"""") &"><td _type=""name""><p align=""left""><b> · "& s.name & "</b>"&IIf(s.description<>"","<br/><span class='note'>&nbsp;&nbsp;"&s.description&"</span>","")&"</p></td><td>"& ExportConfig(name,vbsunescape(value),s) &"</td></tr>"
End Function

Function ExportConfig(s,m,o)
	If o.customhtml<>"" Then
		ExportConfig=o.customhtml
	Else
		If Not o.textarea Then
			ExportConfig="<input type=""text"" name=""" & s & """ id=""" & s & """ value=""" & TransferHTML(m,"[html-format]") & """"
			If o.bool Then
				ExportConfig=ExportConfig & " class=""checkbox""/>"
			Else
				ExportConfig=ExportConfig & " "&o.disabled & " />"
			End If
		Else
			ExportConfig="<textarea cols=""3"" rows=""6"" name=""" & s & """ id=""" & s & """ "&o.disabled&" >"& TransferHTML(m,"[textarea]") & "</textarea>"
	
		End If
	End If
End Function
%>
<script language="javascript" runat="server">
var aryCfgList={"base":[],"database":[],"global":[],"page":[],"build":[],"ubb":[],"test":[],"readonly":[],"others":[]};
function js_getAllConfigs(){
	for (var i in aryCfgList){
		aryCfgList[i].push("<table width=\"100%\"><tr height=\"40\"><th width=\"30%\">配置项</th><th width=\"69%\">配置</th></tr>");
	}
	for(i=1;i<=BlogConfig.Count;i++){	
		var s=getInformation(BlogConfig.Meta.Names(i))	
		aryCfgList[s.type].push(MakeTr(s,BlogConfig.Meta.Names(i),BlogConfig.Meta.Values(i),""))
	}
	for (i in aryCfgList){
		aryCfgList[i].push("</table>");
	}
}
function getInformation(s) {
	var json = {
		"type": "others",
		"description": "",
		"name": s,
		"css": "",
		"disabled": "",
		"textarea": false,
		"bool": false,
		"customhtml" : ""
	};
	switch (s.toUpperCase()) {
	case "ZC_BLOG_HOST":
		json.name = "域名配置";
		json.description = "默认自动读取当前网址,如果系统的识别功能出现问题,或是需要修改网站域名，请先打开“强制启用博客域名”";
		try {
			if (ZC_PERMANENT_DOMAIN_ENABLE != undefined) {
				if (ZC_PERMANENT_DOMAIN_ENABLE) {
					json.disabled = ""
				} else {
					json.disabled = "disabled=\"disabled\""
				}
			} else {
				json.disabled = "disabled=\"disabled\""
			}
		} catch(e) {
			json.disabled = "disabled=\"disabled\""
		}
		json.type = "base";
		break;
	case "ZC_BLOG_TITLE":
		json.name = "设置博客标题";
		json.type = "base";
		break;
	case "ZC_BLOG_SUBTITLE":
		json.name = "设置博客副标题";
		json.type = "base";
		break;
	case "ZC_BLOG_NAME":
		json.name = "博客名";
		json.description = " 在2.0中等同于ZC_BLOG_TITLE，没有任何用处";
		json.disabled = "disabled=\"disabled\"";
		json.css += " throwed ";
		json.type = "base";
		break;
	case "ZC_BLOG_SUB_NAME":
		json.name = "博客描述";
		json.description = " 在2.0中等同于ZC_BLOG_SUBTITLE，没有任何用处";
		json.disabled = "disabled=\"disabled\"";
		json.css += " throwed ";
		json.type = "base";
		break;
	case "ZC_BLOG_THEME":
		json.name = "当前主题";
		json.disabled = "disabled=\"disabled\"";
		json.type = "base";
		break;
	case "ZC_BLOG_CSS":
		json.name = "当前样式";
		json.disabled = "disabled=\"disabled\"";
		json.type = "base";
		break;
	case "ZC_BLOG_COPYRIGHT":
		json.name = "底部版权";
		json.type = "base";
		json.textarea = true;
		break;
	case "ZC_BLOG_MASTER":
		json.name = "博客创始人";
		json.type = "base";
		json.description = "当前管理员：" + BlogUser.FirstName
		break;
	case "ZC_BLOG_LANGUAGE":
		json.name = "博客语言";
		json.type = "base";
		break;
	case "ZC_DATABASE_PATH":
		json.name = "Access数据库路径";
		json.type = "database";
		json.description = "这个页面的配置请勿擅自修改！"
		break;
	case "ZC_MSSQL_ENABLE":
		json.name = "是否使用MSSQL";
		json.type = "database";
		json.bool = true;
		break;
	case "ZC_MSSQL_DATABASE":
		json.name = "MSSQL数据库";
		json.type = "database";
		break;
	case "ZC_MSSQL_USERNAME":
		json.name = "MSSQL用户名";
		json.type = "database";
		break;
	case "ZC_MSSQL_PASSWORD":
		json.name = "MSSQL密码";
		json.type = "database";
		break;
	case "ZC_MSSQL_SERVER":
		json.name = "MSSQL服务器";
		json.type = "database";
		break;
	case "ZC_USING_PLUGIN_LIST":
		json.name = "启用插件列表";
		json.disabled = "disabled=\"disabled\"";
		json.type = "global";
		json.textarea = true;
		break;
	case "ZC_BLOG_CLSID":
		json.name = "CLSID";
		json.disabled = "disabled=\"disabled\"";
		json.type = "global";
		json.description = "系统根据博客路径自动生成，无法修改"
		break;
	case "ZC_TIME_ZONE":
		json.name = "使用者时区";
		json.type = "global";
		json.description = "北京时间：+0800"
		break;
	case "ZC_HOST_TIME_ZONE":
		json.name = "服务器时区";
		json.type = "global";
		json.description = "北京时间：+0800"
		break;
	case "ZC_UPDATE_INFO_URL":
		json.name = "后台公告地址";
		json.type = "global";
		json.disabled = "disabled=\"disabled\"";
		break;
	case "ZC_MULTI_DOMAIN_SUPPORT":
		json.name = "多域名支持";
		json.description = "打开则以相对路径显示博客，不再绑定域名";
		json.bool = true;
		json.type = "test";
		break;
	case "ZC_BLOG_VERSION":
		json.name = "Z-Blog版本";
		json.disabled = "disabled=\"disabled\"";
		json.type = "base";
		json.description = "擅自修改会导致运行错误";
		break;
	case "ZC_COMMENT_TURNOFF":
		json.name = "关闭评论";
		json.type = "page";
		json.bool = true;
		break;
	case "ZC_COMMENT_VERIFY_ENABLE":
		json.name = "打开验证码";
		json.type = "page";
		json.bool = true;
		break;
	case "ZC_COMMENT_REVERSE_ORDER_EXPORT":
		json.name = "评论倒序输出";
		json.type = "page";
		json.bool = true;
		break;
	case "ZC_COMMNET_MAXFLOOR":
		json.name = "评论盖楼最大层数";
		json.description = "管理员后台回复不受限制";
		json.type = "page";
		break;
	case "ZC_VERIFYCODE_STRING":
		json.name = "验证码可用字符";
		json.type = "page";
		json.description = "支持0-9A-Za-z"
		break;
	case "ZC_VERIFYCODE_WIDTH":
		json.name = "验证码宽度";
		json.type = "page";
		break;
	case "ZC_VERIFYCODE_HEIGHT":
		json.name = "验证码高度";
		json.type = "page";
		break;
	case "ZC_DISPLAY_COUNT":
		json.name = "每页显示数量";
		json.type = "page";
		break;
	case "ZC_RSS2_COUNT":
		json.name = "RSS显示数量";
		json.type = "page";
		break;
	case "ZC_SEARCH_COUNT":
		json.name = "搜索结果显示数量";
		json.type = "page";
		break;
	case "ZC_PAGEBAR_COUNT":
		json.name = "分页条显示数量";
		json.type = "page";
		break;
	case "ZC_MUTUALITY_COUNT":
		json.name = "相关文章显示数量";
		json.type = "page";
		break;
	case "ZC_COMMENTS_DISPLAY_COUNT":
		json.name = "评论每页显示数量";
		json.type = "page";
		break;
	case "ZC_USE_NAVIGATE_ARTICLE":
		json.name = "文章页显示上下篇导航";
		json.type = "page";
		json.bool = true;
		break;
	case "ZC_RSS_EXPORT_WHOLE":
		json.name = "输出全文RSS";
		json.type = "page";
		json.bool = true;
		break;
	case "ZC_MANAGE_COUNT":
		json.name = "后台管理数量";
		json.type = "page";
		break;
	case "ZC_REBUILD_FILE_COUNT":
		json.name = "每次重建条数";
		json.type = "build";
		break;
	case "ZC_REBUILD_FILE_INTERVAL":
		json.name = "重建间隔（秒）";
		json.type = "build";
		json.css += " throwed ";
		json.description = "2.0的重建采用批处理，不再有间隔。"
		break;
	case "ZC_UBB_ENABLE":
		json.name = "打开UBB功能";
		json.description = "2.0版已弃用UBB编辑,1.8升级用户打开。";
		json.type = "ubb";
		break;
	case "ZC_UBB_LINK_ENABLE":
		json.name = "——UBB链接转换";
		json.type = "ubb";
		break;
	case "ZC_UBB_FONT_ENABLE":
		json.name = "——UBB字体转换";
		json.type = "ubb";
		break;
	case "ZC_UBB_CODE_ENABLE":
		json.name = "——UBB代码转换";
		json.type = "ubb";
		break;
	case "ZC_UBB_FACE_ENABLE":
		json.name = "——UBB表情转换";
		json.type = "ubb";
		break;
	case "ZC_UBB_IMAGE_ENABLE":
		json.name = "——UBB图片转换";
		json.type = "ubb";
		break;
	case "ZC_UBB_MEDIA_ENABLE":
		json.name = "——UBB多媒体转换";
		json.type = "ubb";
		break;
	case "ZC_UBB_FLASH_ENABLE":
		json.name = "——UBB FLASH转换";
		json.type = "ubb";
		break;
	case "ZC_UBB_TYPESET_ENABLE":
		json.name = "——UBB排版标签转换";
		json.type = "ubb";
		break;
	case "ZC_UBB_AUTOLINK_ENABLE":
		json.name = "——UBB自动链接";
		json.type = "ubb";
		break;
	case "ZC_UBB_AUTOKEY_ENABLE":
		json.name = "——UBB关键字自动替换（弃用）";
		json.disabled = "disabled=\"disabled\"";
		json.css += " throwed ";
		json.type = "ubb";
		break;
	case "ZC_EMOTICONS_FILENAME":
		json.name = "表情文件名（弃用）";
		json.disabled = "disabled=\"disabled\"";
		json.css += " throwed ";
		json.type = "page";
		break;
	case "ZC_EMOTICONS_FILETYPE":
		json.name = "表情图片后缀名";
		json.description = "图片后缀名，用 | 分隔，如png|gif|jpg";
		json.type = "page";
		break;
	case "ZC_EMOTICONS_FILESIZE":
		json.name = "表情大小（弃用）";
		json.disabled = "disabled=\"disabled\"";
		json.css += " throwed ";
		json.type = "page";
		break;
	case "ZC_UPLOAD_FILETYPE":
		json.name = "允许上传文件类型";
		json.type = "global";
		json.description = "以|做为分隔"
		break;
	case "ZC_UPLOAD_FILESIZE":
		json.name = "最大文件大小";
		json.type = "global";
		break;
	case "ZC_UPLOAD_DIRBYMONTH":
		json.name = "上传附件按月存档";
		json.css += " throwed ";
		json.disabled = "disabled=\"disabled\"";
		json.type = "global";
		break;
	case "ZC_UPLOAD_DIRECTORY":
		json.name = "附件保存目录";
		json.type = "global";
		break;
	case "ZC_USERNAME_MIN":
		json.name = "最小用户名长度";
		json.type = "global";
		break;
	case "ZC_USERNAME_MAX":
		json.name = "最大用户名长度";
		json.type = "global";
		break;
	case "ZC_PASSWORD_MIN":
		json.name = "最小密码长度";
		json.type = "global";
		break;
	case "ZC_PASSWORD_MAX":
		json.name = "最大密码长度";
		json.type = "global";
		break;
	case "ZC_EMAIL_MAX":
		json.name = "最大邮件长度";
		json.type = "global";
		break;
	case "ZC_HOMEPAGE_MAX":
		json.name = "最大网址长度";
		json.type = "global";
		break;
	case "ZC_CONTENT_MAX":
		json.name = "最大评论长度";
		json.type = "global";
		break;
	case "ZC_STATIC_TYPE":
		json.name = "静态文件后缀名";
		json.type = "build";
		break;
	case "ZC_STATIC_DIRECTORY":
		json.name = "静态文章存放路径";
		json.type = "build";
		break;
	case "ZC_TEMPLATE_DIRECTORY":
		json.name = "主题模板存放路径";
		json.type = "base";
		json.description = "在这里修改后，必须同时修改当前主题文件夹的TEMPLATE文件夹为在这里定义的文件夹"
		break;
	case "ZC_STATIC_MODE":
		json.name = "列表页模式（动态或伪静态）";
		json.description = "三种模式：ACTIVE   MIX   REWRITE，用于配合静态或伪静态插件";
		json.type = "build";
		break;
	case "ZC_ARTICLE_REGEX":
		json.name = "文章网址格式";
		json.description = "网址格式，下面都是，用于配合静态或伪静态";
		json.type = "build";
		break;
	case "ZC_PAGE_REGEX":
		json.name = "页面网址格式";
		json.type = "build";
		break;
	case "ZC_CATEGORY_REGEX":
		json.name = "分类网址格式";
		json.type = "build";
		break;
	case "ZC_USER_REGEX":
		json.name = "用户网址格式";
		json.type = "build";
		break;
	case "ZC_TAGS_REGEX":
		json.name = "标签网址格式";
		json.type = "build";
		break;
	case "ZC_DATE_REGEX":
		json.name = "日期网址格式";
		json.type = "build";
		break;
	case "ZC_DEFAULT_REGEX":
		json.name = "首页网址格式";
		json.type = "build";
		break;

	case "ZC_DISPLAY_COUNT_WAP":
		json.name = "WAP开关（由模板决定）";
		json.type = "readonly";
		json.disabled = " disabled='disabled'";
		break;
/*
	case "ZC_COMMENT_COUNT_WAP":
		json.name = "WAP单页显示评论数量";
		json.type = "wap";
		break;
	case "ZC_PAGEBAR_COUNT_WAP":
		json.name = "WAP评论条显示条数";
		json.type = "wap";
		break;
	case "ZC_SINGLE_SIZE_WAP":
		json.name = "WAP分页查看文章时单页字数";
		json.type = "wap";
		break;
	case "ZC_SINGLE_PAGEBAR_COUNT_WAP":
		json.name = "WAP文章分页数（未启用）";
		json.type = "wap";
		json.css += " throwed ";
		break;
	case "ZC_FILENAME_WAP":
		json.name = "WAP文件地址";
		json.type = "wap";
		json.css += " throwed ";
		break;
	case "ZC_WAPCOMMENT_ENABLE":
		json.name = "打开WAP评论";
		json.type = "wap";
		json.bool = true;
		break;
	case "ZC_DISPLAY_MODE_ALL_WAP":
		json.name = "WAP显示全文";
		json.type = "wap";
		json.bool = true;
		break;
	case "ZC_DISPLAY_CATE_ALL_WAP":
		json.name = "WAP显示导航";
		json.type = "wap";
		json.bool = true;
		break;
	case "ZC_DISPLAY_PAGEBAR_ALL_WAP":
		json.name = "WAP分页条";
		json.type = "wap";
		json.bool = true;
		break;
	case "ZC_WAP_MUTUALITY_LIMIT":
		json.name = "WAP相关文章数量";
		json.type = "wap";
		break;
*/
	case "ZC_SYNTAXHIGHLIGHTER_ENABLE":
		json.name = "打开代码高亮";
		json.description = "SyntaxHighlighter代码高亮功能";
		json.bool = true;
		json.type = "global";
		break;
	case "ZC_CODEMIRROR_ENABLE":
		json.name = "打开源码编辑高亮";
		json.description = "对应uEditor的CodeMirror，若源码编辑较卡建议禁用";
		json.bool = true;
		json.type = "global";
		break;
	case "ZC_ARTICLE_EXCERPT_MAX":
		json.name = "自动截取摘要字数";
		json.type = "global";
		break;
	case "ZC_UNCATEGORIZED_NAME":
		json.name = "未分类分类名";
		json.type = "global";
		break;
	case "ZC_UNCATEGORIZED_ALIAS":
		json.name = "未分类分类别名";
		json.type = "global";
		break;
	case "ZC_UNCATEGORIZED_COUNT":
		json.name = "未分类分类文章计数";
		json.disabled = " disabled='disabled' ";
		json.type = "global";
		break;
	case "ZC_POST_STATIC_MODE":
		json.name = "文章页模式（动态、静态或伪静态）";
		json.description = "三种模式：ACTIVE   MIX   REWRITE，用于配合静态或伪静态插件";
		json.type = "build";
		break;
	case "ZC_HTTP_LASTMODIFIED":
		json.name = "发送LastModified头";
		json.description = "CDN用户专用，提醒SQUID等程序更新页面缓存";
		json.bool = true;
		json.type = "test";
		break;
	case "ZC_PERMANENT_DOMAIN_ENABLE":
		json.name = "强制锁定博客地址";
		json.description = "若启用，则ZC_BLOG_HOST配置项生效";
		json.type = "base";
		json.bool = true;
		break;
	case "ZC_ARCHIVES_OLD_LISTTYPE":
		json.name = "归档以老样式显示"
		json.bool = true;
		json.type = "page";
		break;
	case "ZC_DEFAULT_PAGES_TEMPLATE":
		json.name = "首页分页指定模板";
		json.description = "如使用default则设为空，否则填写模板文件名（大写）";
		json.type = "page";
		break;
	case "ZC_BLOG_LANGUAGEPACK":
		json.name = "语言包"
		json.description = "部分模板和插件可能依然显示其它的语言。";
		json.type = "base";
		json.customhtml = function(){
			var result='<select id=\"ZC_BLOG_LANGUAGEPACK\" name=\"ZC_BLOG_LANGUAGEPACK\">"';
			var obj,strl,obj2;
			obj2=new Enumerator(new ActiveXObject("Scripting.FileSystemObject").GetFolder(BlogPath+"zb_users\\LANGUAGE").Files);
			for (;!obj2.atEnd();obj2.moveNext()){

				var obj=obj2.item();

				strl=LoadFromFile(obj.Path,"utf-8").split("</language>");
				//Response.Write(LoadFromFile(obj.Path,"utf-8"));Response.End();
				if(strl.length>0){ 	
					result+="<option value=\""+obj.Name.split(".")[0]+"\" "+(ZC_BLOG_LANGUAGEPACK+".asp"==obj.Name?"selected":"")+">";
					result+=TransferHTML(strl[0].split("<language>")[1] + "(" + obj.Name + ")","[html-format]") +"</option>";
				}
			}
			result+="</select>";
			return result;
		}()
	break;
	case "ZC_COMMENT_EXCERPT_MAX":
		json.name = "侧栏评论最大字数";
		json.type = "page";
	break;
	
	default:
		if (/^ZC_SIDEBAR_ORDER\d{0,1}(_OLD)?$/ig.test(s)) {
			var m = /^ZC_SIDEBAR_ORDER(\d{0,1})(_OLD)?$/ig.exec(s);
			json.name = "侧栏" + (m[2] == "" ? "": "原") + "排序 (" + (m[1] == "" ? 1: m[1]) + ")";
			json.disabled = "disabled=\"disabled\"";
			json.type = "global"
		}
	}
	if (json.disabled != "") {
		json.type = "readonly"
	}
	if (json.type == "ubb") {
		json.bool = true;
	}
	return json
}
</script>