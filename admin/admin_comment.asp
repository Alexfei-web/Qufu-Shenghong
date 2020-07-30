<!--#include file="admin_checkuser.asp"-->
<!--#include file="checkpost.asp"-->
<HTML><HEAD><title>管理页面</title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<!--#include file="css/admin_main.css"-->
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
<%
select case Request("action")
	Case "del"
	     Call del()
	Case "show"
	     Call show()	 
	case else
	    Call main()
	End select
        Call CloseDatabase()                                   
%>



<% sub main() %>                                      <!---------------留言列表项------------>
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
 
</TABLE>
<br>

<table width="80%"  border="0" align="center" cellpadding="2" cellspacing="1" Class=TableBorder>
<form action="admin_comment.asp?action=del" name=webmate method="post">
  <tr height="25"> 
    <th width="44%">标题</th>
    <th width="16%">联系人</th>
	<th width="20%">联系电话</th>
    <th width="11%">添加时间</th>
	
	<th width="9%">选择</th>
  </tr>
<%
dim page
dim rs,sql

Set rs= Server.CreateObject("ADODB.Recordset")
sql="select * from comment "
rs.open sql,conn,1,1
if rs.eof then
response.write "<tr><td class=fl1 colspan=""7"">还没有任何留言！</td></tr>" 
End if

 if request("page")<>"" then  
    Page = cint(Request.QueryString("page"))                                 'page值为接受值 
 else
    Page=1
 end if 
const PageSize = 20                                                          '每页显示记录数
rs.PageSize=PageSize

 
    if Page > rs.PageCount then                                              '如果接收的页数大于总页数
        Page = rs.PageCount                                                  '设置当前显示页等于最后页        
    end if 
	if Page <= 0 then                                                        '如果page小于等于0
        Page = 1                                                             '设置PAGE等于第一页
    end if
        rs.AbsolutePage = Page                                               '如果大于零,显示当前页等于接收的页数 
    


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
	<a href="admin_comment.asp?action=show&id=<%=rs("id")%>"><%=rs("title")%></a></td>
    <td class=<%=flcss%> align="center"><%=rs("receiver")%></td>
    
	<td class=<%=flcss%> align="center"><%=rs("phone")%></td>
	<td class=<%=flcss%> align="center"><%=rs("time")%></td>
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
            <a href="admin_comment.asp?page=1">首页</a> | 
            <a href="admin_comment.asp?page=<%= Page-1 %>">上一页</a> | 
            <a href="admin_comment.asp?page=<%= Page+1 %>">下一页</a> |
            <a href="admin_comment.asp?page=<%= rs.pagecount %>">尾页</a>
              
            &nbsp;共<%= rs.RecordCount %>个
            
            &nbsp;&nbsp;<input type="checkbox" name="chkall" onClick="CheckAll(this.form)">全选/取消&nbsp;
            &nbsp;<INPUT TYPE="submit" value="删除" name="action" style="width:40pt;height:22" class="button">&nbsp;&nbsp;</td></tr></table>
        </div>
    </td>
</tr>
<%
    rs.close   
    set rs=nothing
%>
</form>
</table>
<% end sub %>


<% sub del()                                 '_______________删除记录项--------------------
Dim idlist,idArr,uid
dim rs,sql,SucDelNum,i
dim arrUploadFiles,intTemp
if Request("id")="" Then
		ErrMsg=ErrMsg+"<li>操作失败，请选择相关内容进行操作！"
		Call Error(ErrMsg)
		Exit Sub
End if

idlist=Request("id")
idArr=Split(idlist,",")
SucDelNum=0
for i = 0 to ubound(idarr)
	uid=Clng(idarr(i))
	set rs=server.CreateObject("adodb.recordset")
	sql="delete from comment where id="& uid
	rs.open sql,conn,1,3
	
	SucDelNum=SucDelNum+1
next
call adminsuc("成功删除&nbsp;<font color=brown>"&SucDelNum&"</font>&nbsp;条留言！")
end sub %>


<% 
sub show()
dim id
id=trim(request.QueryString("id"))
if id="" or not ISNumeric(id) then
Errmsg=Errmsg+"<br>"+"<li>操作失败，参数丢失！"
Call Error(ErrMsg)
exit sub
end if
id=clng(id)
set rs =server.CreateObject("adodb.recordset")
sql="select * from comment where id="&id
rs.open sql,conn,1,1

if rs.eof then
set rs=nothing
Errmsg=Errmsg+"<br>"+"<li>操作失败，找不到此留言！"
Call Error(ErrMsg)
exit sub
end if
%>
<form action="" name="myform" method=post>
  <table cellpadding="2" cellspacing="1" border="0" width="80%" class=TableBorder align=center>
    <tr> 
      <th colspan=2 height=23>游客留言</th>
    </tr>
    <tr>
      <td colspan="2" Class=fl1 align="center"><table width="90%" height="21" border="0" cellpadding="3" cellspacing="0">
       
		<tr>
          <td>&nbsp;&nbsp;&nbsp;&nbsp;标题：<input name="title" type="text" id="title" value="<%= rs("title") %>" size="50" maxlength="80">
          </td>
        </tr>
		<tr>
          <td>&nbsp;&nbsp;联系人：<input name="receiver" type="text" id="receiver" value="<%= rs("receiver") %>" size="35" maxlength="30">
         </td>
        </tr> 
		<tr>
          <td>联系电话：<input name="phone" type="text" id="phone" value="<%= rs("phone") %>" size="35" maxlength="50">
         </td>
        </tr> 
		<tr>
          <td>&nbsp;&nbsp;&nbsp;&nbsp;传真：<input name="fax" type="text" id="fax" value="<%= rs("fax") %>" size="35" maxlength="50">
         </td>
        </tr> 
		<tr>
          <td>&nbsp;&nbsp;&nbsp;&nbsp;邮箱：<input name="email" type="text" id="email" value="<%= rs("email") %>" size="50" maxlength="80">
         </td>
        </tr> 
       
        <tr>
          <td align="center"><b>留言详细信息</b></td>
        </tr>
        <tr>
          <td>
		  <textarea name="content" id="content" style="width:100%;height:550px;" value=""><%= rs("content") %></textarea>
		  </td>
        </tr>
      </table>
      </td>
    </tr>
    
  </table>
</form> 

<% end sub %>
</BODY></HTML>