<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
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
If CheckPluginState("BaiduRWUD")=False Then Call ShowError(48)
BlogTitle="百度主动推送"

If Request("act")="save" Then
  Call SaveToFile(BaiduRWUD_cfg,Replace(Request("data")," ",""),"utf-8",False)
End If

%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain">
  <div id="ShowBlogHint"><%Call GetBlogHint()%></div>
  <div class="divHeader"><%=BlogTitle%>-<small><a href="javascript:;" onclick="$('#help,#divMain2').slideToggle();" title="查看/隐藏关于">关于</a></small></div><div id="help" style="display:none">置百丈玄冰而崩裂，掷须臾池水而漂摇。<table><td>QQ：</td><td><a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=349467624&site=qq&menu=yes"><img border="0" src="http://wpa.qq.com/pa?p=2:349467624:51" alt="置百丈玄冰而崩裂，掷须臾池水而漂摇。" title="置百丈玄冰而崩裂，掷须臾池水而漂摇。"/></a></td><td>349467624</td></tr><tr><td>Q群：</td><td><a target="_blank" href="http://shang.qq.com/wpa/qunwpa?idkey=c632adb705f6fecbec9a2abad09954b513a7e0131309907715ee06c9d58024c5"><img border="0" src="http://pub.idqqimg.com/wpa/images/group.png" alt="独立博客这个圈！" title="独立博客这个圈！"></a></td><td>101578028</td></tr><tr><td>捐赠：</td><td><a href="http://www.wdssmq.com/guestbook.asp#donate" target="_blank" title="wdssmq@qq.com"><img src="https://img.alipay.com/sys/personalprod/style/mc/btn-index.png" width="120" alt="wdssmq@qq.com"></a></td><td>wdssmq#qq.com</td></tr></table></div>
  <div class="SubMenu"><%=BaiduRWUD_SubMenu(0)%></div>
  <div id="divMain2">
    <form border="1" name="edit" id="edit" method="post" action="main.asp?act=save">
      <p>你需要访问<a href="http://zhanzhang.baidu.com/linksubmit/index" target="_blank">http://zhanzhang.baidu.com/linksubmit/index</a> 获取地址，复制如下图所示区域到插件内。</p>
      <p><textarea style="height:60px;width:100%" name="data" id="data"><%=TransferHTML(LoadFromFile(BaiduRWUD_cfg,"utf-8"),"[textarea]")%></textarea></p>
      <p><input type="submit" class="button" value="提交" name="B1"/></p>
      <%If PublicObjFSO.FileExists(BaiduRWUD_log) Then
        Response.Write LoadFromFile(BaiduRWUD_log,"utf-8")
      Else%>
        <p><img src="index.png"/></p>
      <%End If%>
    </form>
    <script type="text/javascript">
    $("#edit").submit(function(e) {
    var textList = $("#data").val().split("\n");
    for(var i = 0; i < textList.length;i++) {
    if (!/^http:\/\/data.zz.baidu.com\/urls.+token/.test(textList[i])) {
    alert(textList[i] + "不符合规范，请重新填写。");
    return false;
    }
    }
    });
    </script>
  </div>
</div>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%Call System_Terminate()%>
