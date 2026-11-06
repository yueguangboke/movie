<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="..\..\..\c_option.asp" -->
<!-- #include file="..\..\..\..\zb_system\function\c_function.asp" -->
<!-- #include file="..\..\..\..\zb_system\function\c_system_lib.asp" -->
<!-- #include file="..\..\..\..\zb_system\function\c_system_base.asp" -->
<!-- #include file="..\..\..\..\zb_system\function\c_system_event.asp" -->
<!-- #include file="..\..\..\..\zb_system\function\c_system_manage.asp" -->
<!-- #include file="..\..\..\..\zb_system\function\c_system_plugin.asp" -->
<!-- #include file="..\..\..\plugin\p_config.asp" -->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)

BlogTitle="SimplePro主题配置"

Dim strColor
Dim c
Set c = New TConfig
c.Load("SimplePro")
If c.Exists("vesion")=true Then
	strColor=c.read("Color")
End If 

%>
<!--#include file="..\..\..\..\zb_system\admin\admin_header.asp"-->
<script type="text/javascript" src="farbtastic.js"></script>
<link rel="stylesheet" href="farbtastic.css" type="text/css" />
 <script type="text/javascript" charset="utf-8">
  $(document).ready(function() {
    $('#picker').farbtastic('#color');
  });
 </script>
<style>
table input{padding: 0;margin:0.25em 0;}
table input#hdbgph{padding: 2px 5px;}
table .button{padding: 2px 12px 5px 12px; margin: 0.25em 0;}
.tc{border: solid 2px #E1E1E1;width: 50px;height: 23px;float: left;margin: 0.25em;cursor: pointer}
.tc:hover,.active{border: 2px solid #2694E8;}
.upinfo{position: relative;left: 3px;top: -19px;color: white;background: #5EAAE4;width: 190px;height: 23px;display: inline-block;text-align: center;opacity: 0.8;filter: alpha(opacity=80);}
</style>
<!--#include file="..\..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain">
	<div id="ShowBlogHint">
	<%Call GetBlogHint()%>
	</div>
	<div class="divHeader"><%=BlogTitle%></div>
	<div class="SubMenu"></div>
	<div id="divMain2"> 
		<form action="save.asp" method="post">
			<table width="100%" border="1" width="100%" class="tableBorder">
				<tr>
					<th scope="col" height="32" width="150px">颜色配置</th>
					<th scope="col" width="120px">				
					<div style="float:left;margin: 0.25em">推荐颜色：</div>
					</th>
					<th >
						<div id="loadconfig">
						<div class="tc" onclick='$("#color").val("#0099CC");$("#color").css("background-color","#0099CC");' style="background-color:#0099cc"></div>
						<div class="tc" onclick='$("#color").val("#D581A2");$("#color").css("background-color","#D581A2");' style="background-color:#D581A2"></div>
						<div class="tc" onclick='$("#color").val("#CC0000");$("#color").css("background-color","#CC0000");' style="background-color:#CC0000"></div>
						<div class="tc" onclick='$("#color").val("#00ccca");$("#color").css("background-color","#00ccca");' style="background-color:#00ccca"></div>
						<div class="tc" onclick='$("#color").val("#ccaf00");$("#color").css("background-color","#ccaf00");' style="background-color:#ccaf00"></div>
						<div class="tc" onclick='$("#color").val("#18ae04");$("#color").css("background-color","#18ae04");' style="background-color:#18ae04"></div>
						</div>
					</th>
				</tr>
				<tr>
					<td>主题色调</td>
					<td><input type="text" id="color" name="color" value=<%=strColor%> onchange='$(this).style.backgroundColor="#003366"' /></div></td>
					<td><div id="picker"></div></td>
				</tr>
			</table>
			<input name="ok" type="submit" class="button" value="保存配置"/>
		</form>
	</div>
</div>
<!--#include file="..\..\..\..\zb_system\admin\admin_footer.asp"-->
<script type="text/javascript">
ActiveTopMenu("aSimpleProManage");
</script> 

<%Call System_Terminate()%>