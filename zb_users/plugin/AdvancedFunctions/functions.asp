<script language="javascript" runat="server">
var advancedfunctions=function(){
	var THIS=this;
	this.config={};
	this.cls={};
	
	//初始化
	var init=function(){
		THIS.config.sql=(ZC_MSSQL_ENABLE.toString().toLowerCase()=="true"?true:false);
		THIS.config.orig=newClass("TConfig");
		THIS.config.orig.Load("AdvancedFunctions");

		if(THIS.config.orig.Read("version")=="" || typeof(THIS.config.orig.Read("version"))=='undefined'){
			
			THIS.config.orig.Write("version","1.2");       //初始化配置
			
			THIS.config.orig.Write("article_hottest","0"); //阅读最多文章
			THIS.config.orig.Write("article_month","0");   //本月最热文章
			THIS.config.orig.Write("article_year","0");    //本年最热文章
			THIS.config.orig.Write("article_cate","0");    //分类最热文章 
			THIS.config.orig.Write("article_random","0");  //随机文章
			  
			THIS.config.orig.Write("comment_most","0");    //评论最多文章
			THIS.config.orig.Write("comment_month","0");   //本月评论最多
			THIS.config.orig.Write("comment_year","0");    //本年评论最多
			THIS.config.orig.Write("comment_cate","0");    //分类评论最多
			
			
			THIS.config.orig.Write("article_latest","0");  //分类最新文章

			THIS.config.orig.Write("cate_showchild","True");  //显示自分类

			THIS.config.orig.Save();                       //保存配置
		}
		else if(THIS.config.orig.Read("version")=="1.0"){
			THIS.config.orig.Remove("cate_count");
			THIS.config.orig.Write("version","1.2");
			THIS.config.orig.Write("cate_showchild","True");
			THIS.config.orig.Save();
		}
		else if(THIS.config.orig.Read("version")=="1.1"){
			THIS.config.orig.Write("version","1.2");
			THIS.config.orig.Write("cate_showchild","True");
			THIS.config.orig.Save();
		}
		
		//配置读取
		var ary=[new VBArray(THIS.config.orig.Meta.Names).toArray(),new VBArray(THIS.config.orig.Meta.Values).toArray()];
		for(var i in ary[0]){
			THIS.config[ary[0][i]]=ary[1][i];
		}

		if(THIS.config.cate_showchild=="True"){THIS.config.cate_showchild=true}else{THIS.config.cate_showchild=false}
		return true
		
	}();
	
	
	this.save=function(obj){
		if(typeof(obj)!="object") obj={};
		if(!obj.id||!obj.name||!obj.htmlid) return false;
		//Response.Write(obj.id);
		var objfunc;
		GetFunction();
		if(FunctionMetas.GetValue(obj.id)==AdvancedFunctions_Empty){
			objfunc=newClass("TFunction");
			objfunc.ID=0;
			objfunc.Name=obj.name;
			objfunc.FileName=obj.id;
			objfunc.HtmlID=obj.htmlid;
			objfunc.Ftype="ul";
			objfunc.Order=20;
			objfunc.MaxLi=0;
			objfunc.SidebarID=10000;
			objfunc.isSystem=false;
		}
		else{
			objfunc=Functions(FunctionMetas.GetValue(obj.id))
		}
		
		objfunc.Source="plugin_advancedfunctions"
		objfunc.Content=obj.content;
		return objfunc.Save();
	
	}
	
	this.make={
		view:function(strWhere,strOrder,intMax,bType){
			var subtemplate=new Array(intMax),strSQL="";
			var objArticle=null,objRS=null,strSQL="",aryFields=[],aryLoad=[];
			var template="<li><a href=\"$url$\" title=\"$title$\">$title_sort$($nums$)</a></li>"
			
			aryFields=new Array(
								"[log_ID]",
								"[log_CateID]",
								"[log_Title]",
								"[log_Content]",
								"[log_Level]",
								"[log_AuthorID]",
								"[log_PostTime]",
								"[log_Url]",
								"[log_FullUrl]",
								"[log_Type]",
								"[log_CommNums]",
								"[log_ViewNums]",
								"[log_Meta]"
			);
			strSQL="SELECT TOP " + intMax + " " + aryFields.join(",") + " FROM [blog_Article] WHERE [log_Level]=4 AND [log_Type]=0 ";
			strSQL+=(strWhere=="" ? "" : "AND " + strWhere);
			strSQL+=(strOrder=="" ? " ORDER BY [log_"+(bType=="comment"?"Comm":"View")+"Nums] DESC" : strOrder);
			
			//Response.Write(strSQL+"<br/>");

			objRS=objConn.Execute(strSQL);
			
			for(var i=0;i<=intMax;i++){
				if(objRS.EOF){break;}
				aryLoad=new Array(
									objRS("log_ID"),
									"",
									objRS("log_CateID"),
									objRS("log_Title"),
									"",
									"",
									objRS("log_Level"),
									objRS("log_AuthorID"),
									AdvancedFunctions_ReturnTime(objRS("log_PostTime")),
									objRS("log_CommNums"),
									objRS("log_ViewNums"),
									0,
									objRS("log_Url"),
									false,
									"",
									objRS("log_FullUrl"),
									objRS("log_Type"),
									objRS("log_Meta")
				)

				objArticle=NewClass("TArticle");

				if(objArticle.LoadInfoByArray(AdvancedFunctions_ReturnArray(aryLoad))){
					subtemplate[i]=template.replace("$url$",objArticle.Url);
					subtemplate[i]=subtemplate[i].replace("$title$",objArticle.Title);
					subtemplate[i]=subtemplate[i].replace("$title_sort$",objArticle.Title.substr(0,20));
					subtemplate[i]=subtemplate[i].replace("$nums$",(bType=="comment"?objArticle.CommNums:objArticle.ViewNums));
				}
				
				objArticle=null;

				objRS.MoveNext;
			}
			if(i<=0){return ""};
			return subtemplate.join("\n");
		}
	}
	
	this.go=function(type,save){
		if(THIS.config[type]=="0"){return ""}
		if(type=='cate_showchild'||type=='version'){return ""}

		var result={
			id:type,
			name:"",
			htmlid:type,
			content:"",
			where:"",
			max:THIS.config[type],
			bType:"article",
			custom:false,
			order:""
		}
		

		switch (type){
			case "article_hottest":
			
				result.name="阅读最多文章";
				result.order="ORDER BY log_CommNums*50 + sqr(log_ViewNums)*10 - (date()-Log_PostTime)*20 DESC";
				result.bType="article";
				
			break;
			case "article_month":
			
				result.name="本月最热文章";
				result.where="(Month([log_PostTime])=Month("+(this.config.sql?"getdate()":"Now()")+")) AND (Year([log_PostTime])=Year("+(this.config.sql?"getdate()":"Now()")+"))";
				result.bType="article";
				
			break;
			case "article_year":
			
				result.name="本年最热文章";
				result.where="(Year([log_PostTime])=Year("+(this.config.sql?"getdate()":"Now()")+"))";
				result.bType="article";
				
			break;

			case "article_latest":
			
				result.name="分类最新文章";
				result.custom=true;
				result.bType="article";
				
				result.content=function(){
					THIS.getCateTree();

					for(var i=0;i<THIS.treeList.length;i++){

						if(THIS.treeList[i]==null) continue;

						var strSQL="";

						for(var j=0;j<THIS.treeList[i].child.length;j++)
						{
							strSQL+=" OR [log_CateID] = " + THIS.treeList[i].child[j].ID + "  "
						}
						result.content=THIS.make.view(
								"([log_cateid]="+THIS.treeList[i].parent.ID+strSQL+")",
								"ORDER BY log_PostTime DESC",
								result.max,
								"article"
						);

						result.id=result.htmlid="divCate"+THIS.treeList[i].parent.ID+"LatestArticle";
						result.name=THIS.treeList[i].parent.Name+"最新文章";
						if(save){THIS.save(result);}

					}
					return true;
				}()
				
			case "article_cate":
			
				result.name="分类最热文章";
				result.custom=true;
				result.bType="article";
				
				result.content=function(){

					THIS.getCateTree();

					for(var i=0;i<THIS.treeList.length;i++){

						if(THIS.treeList[i]==null) continue;

						var strSQL="";

						for(var j=0;j<THIS.treeList[i].child.length;j++)
						{
							strSQL+=" OR [log_CateID] = " + THIS.treeList[i].child[j].ID + "  "
						}
						result.content=THIS.make.view(
								"([log_cateid]="+THIS.treeList[i].parent.ID+strSQL+")",
								"ORDER BY log_ViewNums DESC",
								result.max,
								"article"
						);

						result.id=result.htmlid="divCate"+THIS.treeList[i].parent.ID+"HottestArticle";
						result.name=THIS.treeList[i].parent.Name+"最热文章";
						if(save){THIS.save(result);}

					}
					return true;
				}()
				
			break;
			case "article_random":
			
				 result.name="随机文章";
				 result.order="ORDER BY " + (this.config.sql?"newid()":"rnd(-([log_id]+\""+Math.random()+"\"))");
				 result.bType="article";

				 
			break;
			case "comment_most":
			
				result.name="评论最多文章";
				result.bType="comment";
				
			break;
			case "comment_month":
			
				result.name="本月评论最多";
				result.where="(Month([log_PostTime])=Month("+(this.config.sql?"getdate()":"Now()")+") AND (Year([log_PostTime])=Year("+(this.config.sql?"getdate()":"Now()")+")))"
				result.bType="comment";
				
			break;
			case "comment_year":
			
				result.name="本年评论最多";
				result.where="(Year([log_PostTime])=Year("+(this.config.sql?"getdate()":"Now()")+"))"
				result.bType="comment";
				
			break;
			case "comment_cate":
			
				result.name="分类评论最多";
				result.bType="comment";
				result.custom=true;
				
				result.content=function(){

					THIS.getCateTree();

					for(var i=0;i<THIS.treeList.length;i++){

						if(THIS.treeList[i]==null) continue;

						var strSQL="";
						for(var j=0;j<THIS.treeList[i].child.length;j++)
						{
							strSQL+=" OR [log_CateID] = " + THIS.treeList[i].child[j].ID + "  "
						}

						result.content=THIS.make.view(
								"([log_cateid]="+THIS.treeList[i].parent.ID+strSQL+")",
								"ORDER BY log_CommNums DESC",
								result.max,
								"comment"
						);

						result.id=result.htmlid="divCate"+THIS.treeList[i].parent.ID+"MostCommentsArticle";
						result.name=THIS.treeList[i].parent.Name+"最多评论";
						if(save){THIS.save(result);}

					}
					return true;
				}()
				
			break;

		}

		if(!result.custom) result.content=THIS.make.view(result.where,result.order,result.max,result.bType);
		if(result.content!=''){
			THIS.save(result);
		}
		return result.content


	}

	this.treeList=null;

	this.getCateTree=function()
	{
			var jsCate=new VBArray(Categorys).toArray();
			
			if(THIS.treeList==null)
			{
				var treeList=[];
				//两次循环再怎么都比数据库多查一遍效率高
						
				for(var i=0;i<jsCate.length;i++){

					if(jsCate[i]==AdvancedFunctions_Empty) continue;

					if (jsCate[i].ParentID==0 || !THIS.config.cate_showchild)
					{

						if(treeList[i]==null)
						{
							treeList[i] = {'parent':jsCate[i],'child':[]};
						}
					}
					else
					{
						if(treeList[jsCate[i].ParentID]!=null)
						{
							treeList[jsCate[i].ParentID].child.push(jsCate[i]);
						}
						else
						{
							treeList[jsCate[i].ParentID] = {'parent':jsCate[i],'child':[jsCate[i]]};
						}
					}
				}
				THIS.treeList=treeList;
			}
		return true;

	}
	
	this.GoAll=function(){
		var w=new VBArray(THIS.config.orig.Meta.Names).toArray();
		for(var i in w){
			if(w[i]!="" && typeof(w[i])!='undefined'){THIS.go(w[i],true);}
		}
	}
	
}
function AdvancedFunctions_Init(){
	return new advancedfunctions();
}

function AdvancedFunctions_ReturnArray(jsarray){return AdvancedFunctions_ReturnArray_VBS(jsarray.join(String.fromCharCode(1)))}
</script>
<%
Function AdvancedFunctions_ReturnArray_VBS(jsarraystr)
	AdvancedFunctions_ReturnArray_VBS=Split(jsarraystr,Chr(1))
End Function
Function AdvancedFunctions_ReturnTime(s)
	AdvancedFunctions_ReturnTime=CStr(CDate(s))
End Function

'****************************************
' AdvancedFunctions 子菜单
'****************************************
Function AdvancedFunctions_SubMenu(id)
	Dim aryName,aryPath,aryFloat,aryInNewWindow,i
	aryName=Array("首页")
	aryPath=Array("main.asp")
	aryFloat=Array("m-left")
	aryInNewWindow=Array(False)
	For i=0 To Ubound(aryName)
		AdvancedFunctions_SubMenu=AdvancedFunctions_SubMenu & MakeSubMenu(aryName(i),aryPath(i),aryFloat(i)&IIf(i=id," m-now",""),aryInNewWindow(i))
	Next
End Function
%>