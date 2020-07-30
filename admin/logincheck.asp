
<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>曲阜圣弘塑胶制品有限公司</title>
</head>
<!--#include file="conn.asp"-->
<!--#include file="md5.asp"-->

<body>


<% 

dim aname,apass,FoundErr,ErrMsg
FoundErr=False

aname=replace(trim(request("admin_username")),"'","")
apass=replace(trim(request("admin_password")),"'","")
safecode=replace(trim(Request("admin_safecode")),"'","")

if len(aname)>20 or len(aname)<3  then
   FoundErr=True
   ErrMsg=ErrMsg&"用户名不对！\n\n"
End if

if len(apass)>20 or len(apass)<3  then
   FoundErr=True
   ErrMsg=ErrMsg&"用户密码不对！\n\n"
End if
if Safecode="" then
	FoundErr=True
	ErrMsg=ErrMsg & "验证码不能为空！\n\n"
end if
if Session("Admin_GetCode")="" then
	FoundErr=True
	ErrMsg=ErrMsg & "你登录时间过长，请重新返回登录页面进行登录。\n\n"
end if
if Safecode<>CStr(Session("Admin_GetCode")) then
	FoundErr=True
	ErrMsg=ErrMsg & "您输入的确认码和系统产生的不一致，请重新输入。\n\n"
end if

if FoundErr=True then
   Call LoginError(ErrMsg)
   Conn.close
   Set Conn=Nothing
else 
     apass=md5(apass)
     dim sql,rs
     sql="select username,password,lastlogin,logins from admin where username='"&aname&"' and password='"&apass&"'"
     set rs=server.createobject("adodb.recordset")
      rs.open sql,conn,1,3
     if rs.BOF and rs.EOF then
          ErrMsg="用户名或是密码错误！"
	      Call LoginError(ErrMsg)
          rs.close
          set rs=Nothing
          conn.close
          set conn=Nothing
          response.End
          
	  else
	      session("admin_username")=aname
		  
		  dim t
		  if rs("logins")="" then
		   t=0
		   else
		    t=rs("logins")
		   end if
		  rs("lastlogin")=now()
		  rs("logins")=t+1
		  rs.update
		  rs.close
	      response.Redirect("useradmin.asp")
	  end if
end if
 
 Sub LoginError(EMsg)
    response.write "<script language='javascript'>" & chr(13)
    response.write "alert('"&EMsg&"');" & Chr(13)
    response.write "window.document.location.href='login.asp';"&Chr(13)
    response.write "</script>" & Chr(13)
    Response.End
End Sub
   %>
</body>
</html>
