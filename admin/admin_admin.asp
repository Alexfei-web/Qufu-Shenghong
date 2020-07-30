<!--#include file="admin_checkuser.asp"-->
<!--#include file="md5.asp"-->
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
select case request("action")  '判断行为
	case "SaveNew"
		Call addadmin()
	case "NewAdmin"
		Call addform()
	case "Delete"
		Call RemoveAdmin()
	case "Modify"
		Call ModifyForm()
	case "ChangeAccount"
		Call SaveEdit()
	case else
		call adminlist()
	End select

Call CloseDatabase()

sub adminlist()
%>
&nbsp; 
<table border=0 width=80% align=center Class=TableBorder>
  <tr>
    <td Class=fl2>
      <B><a name=top>说明</a></B>：<br>
      (1)删除管理员的操作将删除后台管理员帐号，不可恢复。<br>
      (2)要修改管理员的后台登陆帐号请点击相应的用户名。<br>
      </td>
  </tr>
  <tr>
    <td class="fl1"><strong>操作导航：</strong><a href=admin_admin.asp?action=NewAdmin><font color=blue>添加管理员</font></a></td>
  </tr>
</table>
<table width=80% align=center cellspacing=1 cellpadding=5 Class=TableBorder>
  <tr> 
    <th width="18%">用户名</th>
    <th width="25%">管理身份</th>
	<th width="9%">登陆次数</th>
    
    <th width="28%">最后登陆时间</th>
    
    <th width="20%">操作</th>
  </tr>
<%
    dim sql,rs,i,css
    set rs=server.CreateObject("adodb.recordset")
    sql="select username,logins,lastlogin from admin"
    rs.open sql,conn,1,1
    
    if rs.eof then 
%>
<tr><td>还没有管理员！</td></tr>
<%
    else
    i=1
    css="fl1"
    do until rs.eof
    if i mod 2=0 then css="fl2"
%>
  <tr> 
    <td Class=<%=css%> align="center"> 
      <%=rs("username")%>
    </td>
    <td align=center Class=<%=css%>> 系统管理员 </td>
    <td Class=<%=css%> align=center><%=rs("logins")%></td>
    <td align=center Class=<%=css%>> <%=rs("lastlogin")%></td>
    
    <td align=center Class=<%=css%>><a href="admin_admin.asp?action=Delete&username=<%=rs("username")%>" onClick="return confirm('您确定要删除吗？\n\n请保留一位系统管理员，\n否则将无法进入后台！')">删除</a> 
  </tr>
<%
    rs.movenext '移动到数据表下一条记录
    i=i+1
    loop
    end if
    set rs=nothing
%>
</table>
<%
    end sub
    Sub addform()
%>


<BR><BR><BR><BR>
<form action="admin_admin.asp?action=SaveNew" method="post">
<table width="80%" border="0" cellspacing="1" cellpadding="4" align=center Class=TableBorder>
<tr><th colspan=2>添加管理员</th></tr>
<tr><td width=45% align=right Class=fl1>用户名&nbsp;</td><td Class=fl1><input name="username" type="text" size=30>&nbsp;</td>
</tr>
<tr><td width=45% align=right Class=fl2>登陆密码&nbsp;</td><td Class=fl2><input name="userpwd" type="text" size=30></td></tr>
<tr><td colspan=2 align=center Class=fl2><input type="submit" name="Submit" value="&nbsp;添&nbsp;加&nbsp;" class=button>
</td></tr></table>
</form>

<%End Sub%>

<%  
    Sub addadmin() 
    dim sql,rs
	Dim username,userpwd
	username=replace(Request.form("username"),"'","")
	if username="" Then
		ErrMsg=ErrMsg+"<br>"+"<li>操作失败，请填写用户名！"
        Call Error(ErrMsg)
        exit sub
	End if
	
	userpwd=trim(Request.form("userpwd"))
	if userpwd="" Then
		ErrMsg=ErrMsg+"<br>"+"<li>操作失败，请填写登陆密码！"
        Call Error(ErrMsg)
		exit sub
	Else
		userpwd=md5(userpwd)                               '进行加密
	End if
    set rs=server.CreateObject("adodb.recordset")
	sql="select * from admin where username='"&username&"'"
	rs.open sql,conn,1,3
	
	if not (rs.eof and rs.bof) Then
		ErrMsg=ErrMsg+"<br>"+"<li>操作失败，该用户已经是管理员！"
        Call Error(ErrMsg)
		exit sub
	End if
	Set rs=Nothing
	
    set rs=server.CreateObject("adodb.recordset")
	sql="select * from admin"
	rs.open sql,conn,1,3
	rs.addnew
	rs("username")=trim(username)
	rs("password")=userpwd
	rs.update
	rs.close
	response.Redirect("useradmin.asp")
  
    End Sub   
%>
 
 
<% 
    sub RemoveAdmin()
    dim username
    username=replace(trim(request("username")),"'","")
    if username="" then
    ErrMsg=ErrMsg+"<br>"+"<li>操作失败，参数丢失！"
        Call Error(ErrMsg)
		exit sub
    end if
    set rs=server.CreateObject("adodb.recordset")
	sql="select username from admin where username='"&username&"'"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		ErrMsg=ErrMsg+"<br>"+"<li>操作失败！没有找到要进行操作的对象！"
    Call Error(ErrMsg)
	else
	    set rs=server.CreateObject("adodb.recordset")
		sql="delete from admin where username='"&username&"'"
		rs.open sql,conn,1,3
		call adminsuc("删除成功！！！")
		
	end if
	set rs=nothing
    end sub 
%> 
</BODY></HTML>