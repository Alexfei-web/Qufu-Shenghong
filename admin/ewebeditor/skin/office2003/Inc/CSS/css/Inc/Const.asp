<!--#include file="conn.asp"-->
<!--#Include File="pst.asp"-->
<%
Set fl = New fl_Class
if fl.Setup(5)="1" then
Response.Redirect("Error.asp")
response.End
end if
%>