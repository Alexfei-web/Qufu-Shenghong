<!--#include file="conn.asp"-->
<!--#include file="checkpost.asp"-->
<%
    if ChkPost=false then
	'emsg="�벻Ҫ������վ���ύ��"
    response.Redirect("login.asp?emsg=�벻Ҫ������վ���ύ��")
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
<title>��������</title>
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
