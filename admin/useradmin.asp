<!--#include file="conn.asp"-->
<!--#include file="checkpost.asp"-->
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
    else
%>
<html>
<head>
<title>管理中心</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head><frameset rows="675" cols="180,*" border=0 name="bbs"> 
        <frame src="admin_left.asp" noresize name="list" marginwidth="0" marginheight="0">
        <frame src="admin_contact.asp" name="main" marginwidth="0" marginheight="0">
</frameset>
<noframes></noframes>
<body>

</body>
</html>
<%End if%>
<%
Call closedatabase()
%>
