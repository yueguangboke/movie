<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%'UI设计部分有参考coolmud的列表插件%>
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
Call MakeBlogRebuild_Core

%>

<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <!--<div class="SubMenu"></div>-->
          <div id="divMain2"> 
            <script type="text/javascript">ActiveLeftMenu("aPlugInMng");</script>
            <form name="form1" id="form1" action="save.asp?act=save" method="post" enctype="application/x-www-form-urlencoded">
            <p>这是一款侧栏增强插件，能为Z-Blog增加新的侧栏。当然，新的侧栏不仅可以在侧栏调用，它也可以利用标签在其他地方调用。</p>
            <p>强烈建议不用的侧栏条数设置为0，这样可以避免浪费资源。</p>
            <p>插件只负责为您生成侧边栏，剩余的管理还需要您手动通过“侧栏管理”管理。停用插件后，原有侧栏依然可用，但是不再自动更新。</p>
            <table width="100%" style='padding:0px;margin:0px;' cellspacing='0' cellpadding='0' class="tableBorder">
              <tr>
                <th width='20%'><p align="center">侧栏选择</p></th>
                <th width='5%'><p align="center">设置条数</p></th>
                <th><p align="center">说明</p></th>
              </tr>

              <tr>
                <td><b>
                  <label for="m_article_hottest">阅读最多文章</label>
                  </b></td>
                <td><p>
                    <input name="m_article_hottest" type="text" id="m_article_hottest" size="5" value="<%=AdvancedFunctions_.config.article_hottest%>" />
                  </p></td>
                <td>显示访问人数最多的文章。</td>
              </tr>
              <tr>
                <td><b>
                  <label for="m_article_month">本月最热文章</label>
                  </b></td>
                <td><p>
                    <input name="m_article_month" type="text" id="m_article_month" size="5" value="<%=AdvancedFunctions_.config.article_month%>" />
                  </p></td>
                <td>显示30天内访问人数最多的文章。</td>
              </tr>
              <tr>
                <td><b>
                  <label for="m_article_year">本年最热文章</label>
                  </b></td>
                <td><p>
                    <input name="m_article_year" type="text" id="m_article_year" size="5" value="<%=AdvancedFunctions_.config.article_year%>" />
                  </p></td>
                <td>显示365天内访问人数最多的文章。</td>
              </tr>
              <tr>
                <td><b>
                  <label for="m_comment_most">评论最多文章</label>
                  </b></td>
                <td><p>
                    <input name="m_comment_most" type="text" id="m_comment_most" size="5" value="<%=AdvancedFunctions_.config.comment_most%>" />
                  </p></td>
                <td>显示评论最多的文章</td>
              </tr>
              <tr>
                <td><b>
                  <label for="m_comment_month">本月评论最多</label>
                  </b></td>
                <td><p>
                    <input name="m_comment_month" type="text" id="m_comment_month" size="5" value="<%=AdvancedFunctions_.config.comment_month%>" />
                  </p></td>
                <td>显示30天内评论最多的文章</td>
              </tr>
              <tr>
                <td><b>
                  <label for="m_comment_year">本年评论最多</label>
                  </b></td>
                <td><p>
                    <input name="m_comment_year" type="text" id="m_comment_year" size="5" value="<%=AdvancedFunctions_.config.comment_year%>" />
                  </p></td>
                <td>显示365天内评论最多的文章</td>
              </tr>


              <tr>
                <td><b>
                  <label for="m_article_random">随机文章</label>
                  </b></td>
                <td><p>
                    <input name="m_article_random" type="text" id="m_article_random" size="5" value="<%=AdvancedFunctions_.config.article_random%>" />
                  </p></td>
                <td>为了提高效率，只有重建缓存时随机文章才会被更新</td>
              </tr>
              <tr>
                <td>
                  <b><label for="s_cate_showchild">显示子分类</label>
                  </b></td>
                </td>
                <td>
                  <input name="s_cate_showchild" type="text" id="s_cate_showchild" size="5" class="checkbox" value="<%=AdvancedFunctions_.config.cate_showchild%>" />
                </td>
                <td>
                    是否在分类文章侧栏中显示子分类的文章。
                </td>
              </tr>
              <tr>
                <td><b>
                  <label for="m_article_latest">分类最新文章</label>
                  </b></td>
                <td><p>
                    <input name="m_article_latest" type="text" id="m_article_latest" size="5" value="<%=AdvancedFunctions_.config.article_latest%>" />
                  </p></td>
                <td>显示某个分类最新的文章</td>
              </tr>
              <tr>
                <td><b>
                  <label for="m_article_cate">分类最热文章</label>
                  </b></td>
                <td><p>
                    <input name="m_article_cate" type="text" id="m_article_cate" size="5" value="<%=AdvancedFunctions_.config.article_cate%>" />
                  </p></td>
                <td>显示某个分类访问人数最多的文章</td>
              </tr>

              <tr>
                <td><b>
                  <label for="m_comment_cate">分类评论最多</label>
                  </b></td>
                <td><p>
                    <input name="m_comment_cate" type="text" id="m_comment_cate" size="5" value="<%=AdvancedFunctions_.config.comment_cate%>" />
                  </p></td>
                <td>分类评论最多的文章</td>
              </tr>

            </table>
            <input type="submit" value="保存" class="button"/>
            </form>
          </div>
        </div>
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->