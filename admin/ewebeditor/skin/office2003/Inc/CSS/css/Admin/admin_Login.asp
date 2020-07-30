<!--#include file="ser.asp"-->
<!--#include file="../Inc/md5.asp"-->
<%Response.Expires = -1%>
<html>
<head>
<title><%=fl.Setup(0)%>--css</title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<!--#include file="images/admin_main.css"-->
</hend>
<%
dim adminlogins
adminlogins=9
if session(fl.fl_sn&"_admin_logins")="" then session(fl.fl_sn&"_admin_logins")=0
select case Request("Action")
case "chkLogin"
	Call chklogin()
case "logout"
	Call logout()
case Else
	Call main()
End select
set fl=nothing
Call CloseDatabase()
Sub chklogin()

	dim username,userpwd
	dim Adminsql,AdminRs
    if session(fl.fl_sn&"_admin_logins")>=adminlogins then
    Response.Write("<script language=JavaScript>alert('!');document.location.href='../';</script>")
    exit sub
    end if
	
	username=fl.CheckStr(trim(Request.form("username")))
	userpwd=fl.CheckStr(trim(Request.form("password")))
	if username="" or userpwd="" Then
		Errmsg=Errmsg+"<br>"+"<li>。"
		Call Error(ErrMsg)
	ElseIf isValidstring(username)<>"" Then
		ErrMsg=ErrMsg+"<br>"+"<li>！"
		Call Error(ErrMsg)
	Elseif fl.CodeIsTrue()=false then
	Errmsg=Errmsg+"<br>"+"<li>。</b>"
		Call Error(ErrMsg)
	else
		userpwd=md5(userpwd,32)       'MD5  32位加密
		Adminsql="select * from Fl_admin where UserName='"&username&"' and UserPwd='"&userpwd&"'"
		set AdminRs=fl.execute(Adminsql)
		if AdminRs.eof and AdminRs.bof Then
		    session(fl.fl_sn&"_admin_logins")=session(fl.fl_sn&"_admin_logins")+1
			Errmsg=Errmsg+"<br>"+"<li>！"&_
			"<br><li><font color=red>！</font>"&_
			"<br><li>请<a href=admin_login.asp target=_top><font color=red>重新登录</font></a>！"
		Call Error(ErrMsg)
		elseif AdminRs("IsLocked")=1 then
		Errmsg=Errmsg+"<br>"+"<li>！"
			Call Error(ErrMsg)
		else
			fl.execute("update Fl_Admin set LastLogin="&SqlNowString&",LastIP='"&fl.UserTrueIP&"',Logins=Logins+1 where UserName='"&username&"'")
		   Session(fl.fl_sn&"_AdminFlag")=AdminRs("UserFlag")
		   Session(fl.fl_sn&"_AdminName")=AdminRs("UserName")
		   Session(fl.fl_sn&"_AdminType")=AdminRs("UserType")
		   session.timeout=40
		   response.redirect "Index.asp"
		end if
		Set AdminRs=Nothing
	end if

End Sub


Sub logout()
    fl.execute("update Fl_Admin set LastLogout="&SqlNowString&" where UserName='"&AdminName&"'")
	Session(fl.fl_sn&"_AdminFlag")=""
	Session(fl.fl_sn&"_AdminName")=""
	Session(fl.fl_sn&"_AdminType")=""
	response.redirect "../Index.asp"
End Sub

Sub main()
%>
<BR><BR><BR><BR><BR><BR>
<%
if session(fl.fl_sn&"_admin_logins")>=adminlogins then
Response.Write("<script language=JavaScript>alert('!');document.location.href='../';</script>")
Response.End
else
%>
<script language=javascript>
<!--
function CheckForm()
{
	if(document.Login.username.value=="")
	{
		alert("！");
		document.Login.username.focus();
		return false;
	}
	if(document.Login.password.value == "")
	{
		alert("！");
		document.Login.password.focus();
		return false;
	}
	if (document.Login.CodeStr.value==""){
       alert ("！");
       document.Login.CodeStr.focus();
       return(false);
    }
}

//-->
</script>
  <form name="Login" action="admin_Login.asp" method="post" onSubmit="return CheckForm();">
    <input type=hidden name=Action value=chkLogin>
<table cellpadding="1" cellspacing="0" border="0" align=center style="border: outset 3px;width:0;">
<tr>
	<td>
	<table style="width:500" border=0 cellspacing=0 cellpadding=3 align=center class=tablefoot>
		<tr>
		  <th valign=middle colspan=2 height=25>css样式登录</th>
		</tr>
	</table>
	<table style="width:500" border=0 cellspacing=0 cellpadding=3 align=center>
	<tr>
	<td valign=middle colspan=2 align=center class=fl2 height=4></td>
	</tr>
	<tr>
	<td valign=middle class=fl1 width="30%" align=right><b>css：</b></td>
	<td valign=middle class=fl1><Input name=username type=text onMouseOver="this.style.background='#D6DFF7';this.focus();" onMouseOut="this.style.background='#FFFFFF'" onFocus="this.select(); " style="width:130px"></td></tr>
	<tr>
	<td valign=middle class=fl1 align=right><b>csspass：</b></font></td>
	<td valign=middle class=fl1><Input name=password type=password onMouseOver="this.style.background='#D6DFF7';this.focus();" onMouseOut="this.style.background='#FFFFFF'" onFocus="this.select(); " style="width:130px"></td></tr>
	<tr>
	<td valign=middle class=fl1 align=right><b>css-!：</b></td>
	<td valign=middle class=fl1><Input name=CodeStr type=text style="width:38px" onFocus="this.select(); " onMouseOver="this.style.background='#D6DFF7';this.focus();" onMouseOut="this.style.background='#FFFFFF'" maxlength="4">&nbsp;请 <%=fl.GetCode()%></td></tr>
	<tr>
	<td valign=middle colspan=2 align=center class=fl2><input class=button type=submit name=submit value="登 录"></td>
	</tr>
	</table>
	</td>
</tr>
</table>
</FORM>
<%end if
End Sub

function isValidstring(para)
	on error resume next
	dim str
	dim l,i,invalidchar
	invalidchar="=%?#&;,'+<>()-:\*!/|"&chr(32)&chr(34)&chr(9)
	if isNUll(para) then 
		isValidstring=""
		exit function
	end if
	str=cstr(para)
	if trim(str)="" then
		isValidstring=""
		exit function
	end if
	l=len(str)
	for i = 1 to l
		c = Mid(str, i, 1)
		if InStr(invalidchar,c)>0 then
			isValidstring = c
			exit function
		end if
	next
	isValidstring=""
	if err.number<>0 then err.clear
end function
%>
</html>