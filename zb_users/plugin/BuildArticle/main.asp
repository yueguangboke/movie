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
If CheckPluginState("BuildArticle")=False Then Call ShowError(48)
BlogTitle="分段重建"

If ZC_POST_STATIC_MODE<>"STATIC" Then
	SetBlogHint_Custom "您当前的文章页模式不是静态页，无需文件重建。"
End If
'Call BuildArticle_BuildArticle(1,False)
'Response.End
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<style type="text/css">
tr {
	height: 32px
}
#show_result{
	cursor:wait
}
</style>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"><%=BuildArticle_SubMenu(0)%></div>
          <div id="divMain2"> 
            <script type="text/javascript">ActiveTopMenu("aPlugInMng");</script>
            <%ExportPage%>
          </div>
        </div>
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%Call System_Terminate()%>
<%
Sub ExportPage
%>
<script type="text/javascript">
var data={
	max:<%=objConn.Execute("SELECT MAX(log_ID) FROM blog_Article")(0)%>,
	count:<%=objConn.Execute("SELECT COUNT(log_ID) FROM blog_Article")(0)%>,
	min:<%=objConn.Execute("SELECT MIN(log_ID) FROM blog_Article")(0)%>,
	config:{
		rebuild_count:<%=ZC_REBUILD_FILE_COUNT%>,
		maxtime:0,
		delay:0
	},
	status:{}
}
</script>
<div id="show_result" style="display:none">
  <p>进度：<span id="progress">0</span>%，已用时<span id="second">0</span>秒</p>
  <p id="_jqui_progress"></p>
  <ul id="result_ul">
  </ul>
</div>
<div id="user_input">
  <p>您当前的文章数量：<span id="count">请稍等</span>篇。请选择重建方法</p>
  <form id="form1" name="form1" method="post" action="">
    <ul>
      <li>
        <input type="radio" id="rebuild_type" name="rebuild_type" value="1" checked="checked" />
        <label for="rebuild_type">重建全部文章</label>
      </li>
      <li>
        <input type="radio" id="rebuild_type2" name="rebuild_type" value="2" />
        <label for="rebuild_type2">重建ID从
          <input name="id_min" id="id_min" type="number" />
          到
          <input name="id_max" id="id_max" type="number" />
          的文章</label>
      </li>
      <li>
        <input type="radio" id="rebuild_type3" name="rebuild_type" value="3" />
        <label for="rebuild_type3">重建分类
          <select name="cate_id" id="cate_id">
            <%Dim Category:For Each Category In Categorys:If IsObject(Category) Then%>
            <option value="<%=Category.ID%>"><%=Category.ID%>.<%=Category.Name%></option>
            <%End If:Next%>
          </select>
          的文章</label>
      </li>
      <li>
        <input type="radio" id="rebuild_type4" name="rebuild_type" value="4" />
        <label for="rebuild_type4">重建作者
          <select name="user_id" id="user_id">
            <%Dim User:Call GetUser:For Each User In Users:If IsObject(User) Then%>
            <option value="<%=User.ID%>"><%=User.ID%>.<%=User.Name%></option>
            <%End If:Next%>
          </select>
          的文章</label>
      </li>
      <li>
        <input type="radio" id="rebuild_type5" name="rebuild_type" value="5" />
        <label for="rebuild_type5">重建首页和侧栏</label>
      </li>
      <li>
        <input name="rebuild_stat" type="checkbox" value="1" checked="checked" id="rebuild_stat" />
        <label for="rebuild_stat">重建统计(如速度太慢请取消勾选)</label>
      </li>
      <li>
        <label for="split">重建间隔时间(ms)</label>
        <input name="split" type="number" min="0" max="60000" step="1000" value="0" id="split" />
      </li>
    </ul>
    <p>
      <input name="" type="submit" class="button" value="开始重建" />
    </p>
  </form>
</div>
<script type="text/javascript">
var run=function(){
	$.ajax({
		url:"post.asp",
		type:"POST",
		data:data.status,
		dataType:"json",
		error: function(xhr){
			$("#result_ul").append(function(){
				return "<li>【"+getD()+"】文件重建失败，出现"+xhr.status+"错误，具体错误为"+xhr.responseText+"</li>"
			});
		},
		success: function(result){
			data.status=result;
			$("#result_ul").append(function(){
				return "<li>【"+getD()+"】"+data.status.msg+"</li>"
			});
			//if(data.status.dbmax){
			$("#_jqui_progress").progressbar({value:((data.status.page-1)*data.config.rebuild_count)/data.status.max*100});
			$("#progress").html(parseInt(((data.status.page-1)*data.config.rebuild_count)/data.status.max*100));
			//}
			data.config.maxtime+=parseInt(data.status.time);
			$("#second").html(data.config.maxtime/1000);
			if(data.status.ok){
				$("#_jqui_progress").progressbar({value:100});
				$("#progress").html(100);
				$("#show_result").css("cursor","auto");
				return ""
			}else{
				if(data.config.delay==0){
					return run()
				}
				else{
					$("#result_ul").append("<li>【"+getD()+"】等待"+(data.config.delay/1000)+"秒后继续执行...<\/li>");
					setTimeout("run()",data.config.delay);
					return ""
				}
			}
		}
	})

}
$(document).ready(function(){
	$("#count").html(data.count);
	$("#id_min").attr("min",data.min).attr("max",data.max).val(data.min);
	$("#id_max").attr("min",data.min).attr("max",data.max).val(data.max);
	$("#form1").submit(function(){
		data.config.delay=parseInt($("#split").val());
		data.status={};
		data.status.ok=false;
		data.status.type=parseInt($("[name='rebuild_type']:checked").val());
		data.status.stat=($("#rebuild_stat").attr("checked")?true:false);
		data.status.page=1;
		$("#_jqui_progress").progressbar({min:0,max:100})
		$("#show_result").show();
		$("#user_input").hide();
		switch (data.status.type){
			case 1:
				data.status.type=2;
				data.status.min=data.min;
				data.status.max=data.max;
			case 2:
				data.status.min=parseInt($("#id_min").val());
				data.status.max=parseInt($("#id_max").val());
			break;
			case 3:
				data.status.cate_id=parseInt($("#cate_id").val());
				data.status.min=data.min;
				data.status.max=data.max;
			break;
			case 4:
				data.status.auth_id=parseInt($("#user_id").val());
				data.status.min=data.min;
				data.status.max=data.max;
			break;
		}
		run();
		return false
	})
	<%=IIf(Request.QueryString("build")="all","$('#form1').submit()","")%>
})

function getD(){
	var o=new Date();
	return o.getFullYear()+"-"+(o.getMonth()+1)+"-"+o.getDate()+" "+o.getHours()+":"+o.getMinutes()+":"+o.getSeconds()
}

</script>
<%
End Sub
%>