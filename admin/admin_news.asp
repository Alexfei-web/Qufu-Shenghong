<!--#include file="admin_checkuser.asp"-->
<!--#include file="checkpost.asp"-->
<script type="text/javascript" src="kindeditor/kindeditor.js"></script> 
<script type="text/javascript" src="kindeditor/lang/zh_CN.js"></script> 
<script type="text/javascript">

		KindEditor.ready(function(K) {
			var editor1 = K.create('#editor_id',{
			 width : "650px",
             height : "500px", 
             filterMode : false, 
             resizeMode :1 ,
			   
				cssPath : 'KindEditor/plugins/code/prettify.css',
				uploadJson : 'KindEditor/asp/upload_json.asp',
				fileManagerJson : 'KindEditor/asp/file_manager_json.asp',
				allowFileManager : true,
				afterCreate : function() {
					var self = this;
					K.ctrl(document, 13, function() {
						self.sync();
						K('form[name=myform]')[0].submit();
					});
					K.ctrl(self.edit.doc, 13, function() {
						self.sync();
						K('form[name=myform]')[0].submit();
					});
				}
			});
			
		});
	</script>
<HTML><HEAD><title>管理页面</title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<!--#include file="css/admin_main.css"-->
<style type="text/css">textarea{ word-break:break-all; resizeMode:0;}</style>
</HEAD>
<%
if ChkPost=false then
	'emsg="请不要从其它站点提交表"
    response.Redirect("login.asp?emsg=请不要从其它站点提交表")
    Response.End()
    End if
    dim AdminName
    AdminName=session("admin_username")
    if AdminName="" then
    response.redirect"login.asp"
    end if
%>
<BODY leftMargin=0 topMargin=0 marginwidth="0" marginheight="0">
<script language = "JavaScript">
    function checkadd(){
        if (document.myform.title.value == '') {
            alert('新闻标题不能为空');
		    document.myform.title.focus();
		    return false
		}		
	}
</script>

<%

select case Request("action")
	case "Save"
		Call Save()
	case "add"
		call add()
	Case "edit"
	     Call Edit()
	Case "EditSave"
	     Call EditSave()
	Case "passed"
	      Call passed()
	Case "del"
	     Call del()
	case else
	    Call main()
	End select


Call CloseDatabase()
Sub add()                                            '——————————————添加新闻————————————————


%>
<form action="admin_news.asp?action=Save" name="myform" method=post onSubmit="return checkadd();">
  <table cellpadding="2" cellspacing="1" border="0" width="80%" class=TableBorder align=center>
    <tr> 
      <th colspan=2 height=23>新闻添加</th>
    </tr>
    <tr>
      <td colspan="2" Class=fl1 align="center"><table width="90%" height="21" border="0" cellpadding="3" cellspacing="0">
        <tr>
          <td>&nbsp;新闻标题：<input name="title" type="text" id="title" value="" size="60" maxlength="150">&nbsp;* </td>
        </tr>
       
        <tr>
          <td>&nbsp;新闻来源：<input name="author" type="text" id="author" value="" size="40" maxlength="30"></td>
        </tr>
        <tr>
          <td>&nbsp;新闻关键字：<input name="key" type="text" id="key" size="45" maxlength="500">多个请用，分隔</td>
        </tr>
        <tr>
          <td align="center"><textarea name="content" id="editor_id" style="width:70%;height:550px;visibility:hidden;"></textarea>
		 </td>
        </tr>
       
      </table>
      </td>
    </tr>
    
    <tr align="center"> 
      <td colspan="2" Class=fl2> <input type="submit" name="submit" value="发 表" class=button> 
        &nbsp;&nbsp; <input type=button value="返 回" onClick="location.href='admin_news.asp'" class=button>      </td>
    </tr>
  </table>
</form>
<%end sub%>

<% 
    Sub Save()                                    '添加新闻后   保存
    dim rs,sql,founderr
    founderr=false
    dim Title,Content,Author,key
    
    Title=trim(request.form("title"))
    Author=trim(request.form("author"))
    Content=trim(request.form("content"))
    key=trim(request.form("key"))
    
    if Title="" then
    		founderr=true
    		ErrMsg=ErrMsg & "<br><li>新闻标题不能为空</li>"
    end if
    if Content="" then
    		founderr=true
    		ErrMsg=ErrMsg & "<br><li>新闻内容不能为空</li>"
    end if
    
    
    if founderr=true then
    Call Error(ErrMsg)
    exit sub
    end if
    
    set rs=server.createobject("adodb.recordset")
    sql="select * from news"
    rs.open sql,conn,1,3
    rs.addnew
    rs("title")=Title
    rs("author")=Author
    rs("time")=year(now) & "-" & month(now)  &"- " & day(now)
    rs("content")=Content
    rs("hits")=0
    rs("key")=key
    
    rs.update
    rs.close
    set rs=nothing
    
    call adminsuc("新闻添加成功！")
    End Sub
%>

<% 
    Sub EditSave()                                 '保存编辑新闻内容
    dim rs,sql,founderr,id
    dim Title,Content,Author
    
    Title=trim(request.form("title"))
    Author=trim(request.form("author"))
    
    Content=trim(request.form("content"))
    
    key=trim(request.form("key"))
    
    id =request.form("id")
    
    id=clng(id)
    if Title="" then
    		founderr=true
    		ErrMsg=ErrMsg & "<br><li>新闻标题不能为空</li>"
    end if
    if Content="" then
    		founderr=true
    		ErrMsg=ErrMsg & "<br><li>新闻内容不能为空</li>"
    end if
    
    
    if founderr=true then
    Call Error(ErrMsg)
    exit sub
    end if
    
    set rs=server.createobject("adodb.recordset")
    
    sql="select * from news where id="& id
    rs.open sql,conn,1,3
    if rs.eof then
    rs.close
    set rs=nothing
    ErrMsg=ErrMsg & "<br><li>找不到需要修改的记录</li>"
    Call Error(ErrMsg)
    exit sub
    end if
    rs("title")=Title
    rs("author")=Author
    
    rs("content")=Content
    rs("key")=key
    rs.update
    rs.close
    set rs=nothing
    
    call adminsuc2("新闻修改成功！","admin_news.asp")
    End Sub 
%>

<% sub main() %>                                      <!---------------新闻列表项------------>
<script language="javascript">
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.type.toLowerCase()=="checkbox"&&e.name != 'chkall')
	    e.checked = form.chkall.checked; 
   }
}
</script>
<TABLE border=0 width="80%" align=center Class=TableBorder>
  <TR> 
    <TD class=fl2 colspan="2"><b>说明：</b><br>
      (1)一旦删除记录，就无法恢复，请慎重考虑！</TD>
  </TR>
  <TR> 
    <TD class=fl1>

<a href="admin_news.asp?action=add"><font color=blue>&nbsp;&nbsp; 添加新闻</font></a></TD>
	
  </TR>
</TABLE>
<br>

<table width="80%"  border="0" align="center" cellpadding="2" cellspacing="1" Class=TableBorder>
<form action="admin_news.asp?action=del" name=webmate method="post">
  <tr height="25"> 
    <th width="44%">标题</th>
    <th width="20%">来源</th>
	<th width="16%">添加时间</th>
    <th width="11%">点击次数</th>
	
	<th width="9%">选择</th>
  </tr>
<%
    dim page
    dim rs,sql
    
    Set rs= Server.CreateObject("ADODB.Recordset")
    sql="select * from news "
    rs.open sql,conn,1,1
    if rs.eof then
    response.write "<tr><td class=fl1 colspan=""7"">还没有任何新闻！</td></tr>" 
    End if
    
    if request("page")<>"" then  
        Page = cint(Request.QueryString("page"))                               'page值为接受值 
    else
        Page=1
    end if 
    const PageSize = 20                                                        '每页显示记录数
    rs.PageSize=PageSize 
        if Page > rs.PageCount then                                            '如果接收的页数大于总页数
            Page = rs.PageCount                                                '设置当前显示页等于最后页        
        end if 
    	if Page <= 0 then                                                      '如果page小于等于0
            Page = 1                                                           '设置PAGE等于第一页
        end if
            rs.AbsolutePage = Page                                             '如果大于零,显示当前页等于接收的页数 
    dim i,flcss
    i = 0
    do while not rs.eof
    if i mod 2=0 then
    flcss="fl1"
    else
    flcss="fl2"
    end if
%>
  <tr> 
    <td class=<%=flcss%>>
	<a href="?action=edit&id=<%=rs("id")%>"><%=rs("title")%></a></td>
    <td class=<%=flcss%> align="center"><%=rs("author")%></td>
    <td class=<%=flcss%> align="center"><%=rs("time")%></td>
	<td class=<%=flcss%> align="center"><%= rs("hits") %></td>
	
	<td class=<%=flcss%> align="center"><input name="id" type="checkbox" id="id" value="<%=rs("id")%>"></td>
  </tr>
  <%
          i=i+1
	      if i>=PageSize then exit do
	      rs.movenext
	loop
  %>
  <tr> 
    <td colspan="8" Class=fl2> 
      <div align="right"> 
   
 <table width=""100%""><tr><td align=right>
 <tr><td>
<a href="admin_news.asp?page=1">首页</a> | 
<a href="admin_news.asp?page=<%= Page-1 %>">上一页</a> | 
<a href="admin_news.asp?page=<%= Page+1 %>">下一页</a> |
<a href="admin_news.asp?page=<%= rs.pagecount %>">尾页</a>
  
&nbsp;共<%= rs.RecordCount %>个

&nbsp;&nbsp;<input type="checkbox" name="chkall" onClick="CheckAll(this.form)">全选/取消&nbsp;
&nbsp;<INPUT TYPE="submit" value="删除" name="action" style="width:40pt;height:22" class="button">&nbsp;&nbsp;</td></tr></table>
</div></td>
  </tr>
<%
    rs.close   
    set rs=nothing
%>
</form>
</table>
<% end sub %>


<% 
    sub del()                                 '删除记录项
    Dim idlist,idArr,uid
    dim rs,sql,SucDelNum,i
    dim arrUploadFiles,intTemp
    if Request("id")="" Then
		ErrMsg=ErrMsg+"<li>操作失败，请选择相关新闻进行操作！"
		Call Error(ErrMsg)
		Exit Sub
    End if

    idlist=Request("id")
    idArr=Split(idlist,",")
    SucDelNum=0
    for i = 0 to ubound(idarr)
	uid=Clng(idarr(i))
	set rs=server.CreateObject("adodb.recordset")
	sql="delete from news where id="& uid
	rs.open sql,conn,1,3
	
	SucDelNum=SucDelNum+1
    next
    call adminsuc("成功删除&nbsp;<font color=brown>"&SucDelNum&"</font>&nbsp;条新闻！")
    end sub 
%>



<% 
    Sub Edit()                              '编辑新闻项
    dim id,sql,rs
    id=trim(request("id"))
    if id="" or not ISNumeric(id) then
    Errmsg=Errmsg+"<br>"+"<li>操作失败，参数丢失！"
    Call Error(ErrMsg)
    exit sub
    end if
    id=clng(id)
    set rs =server.CreateObject("adodb.recordset")
    sql="select * from news where id="&id
    rs.open sql,conn,1,1
    
    if rs.eof then
    set rs=nothing
    Errmsg=Errmsg+"<br>"+"<li>操作失败，找不到此新闻！"
    Call Error(ErrMsg)
    exit sub
    end if
%>
<form action="admin_news.asp?action=EditSave" name="myform" method=post onSubmit="return checkadd();">
  <table cellpadding="2" cellspacing="1" border="0" width="80%" class=TableBorder align=center>
    <tr> 
      <th colspan=2 height=23>新闻修改</th>
    </tr>
    <tr>
      <td colspan="2" Class=fl1 align="center"><table width="90%" height="21" border="0" cellpadding="3" cellspacing="0">
        <tr>
          <td>&nbsp;新闻标题：<input name="title" type="text" id="Title" value="<%=rs("title")%>" size="60" maxlength="150">&nbsp;* </td>
        </tr>
       
        <tr>
          <td>&nbsp;新闻来源：<input name="author" type="text" id="Author" value="<%=rs("author")%>" size="20" maxlength="30">
		  <input name="id" type="hidden" id="id" value="<%=rs("id")%>"></td></tr>
            
        
        <tr>
          <td>&nbsp;新闻关键字：<input name="key" type="text" id="key" size="40" maxlength="500" value="<%=rs("key")%>"> 
          多个请用 , 分隔</td>
        </tr>
        <tr>
          <td align="center"><textarea name="content" id="editor_id" style="width:100%;height:550px;visibility:hidden;"><%=rs("content")%></textarea>
		  </td>
        </tr>
    
      </table>
      </td>
    </tr>
    
    <tr align="center"> 
      <td colspan="2" Class=fl2> <input type="submit" name="Submit" value="修 改" class=button> 
        &nbsp;&nbsp; <input type=button value="返 回" onClick="location.href='admin_news.asp'" class=button>      </td>
    </tr>
  </table>
</form>
<%
set rs=nothing
end sub
%>

</BODY></HTML>