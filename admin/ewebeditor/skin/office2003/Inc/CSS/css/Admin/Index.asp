<!--#INCLUDE FILE="ser.asp" -->
<%if AdminName="" then
response.redirect"admin_Login.asp"
else
%>
<html>
<head>
<title><%=fl.Setup(0)%>--css</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<frameset rows="675" cols="180,*" border=0 name="bbs"> 
        <frame src="admin_left.asp" noresize name="list" marginwidth="0" marginheight="0">
        <frame src="admin_Main.asp" name="main" marginwidth="0" marginheight="0">
</frameset>
<noframes></noframes>
</body>
</html>
<%End if%>
<%
set fl=nothing
Call CloseDatabase()
%>
