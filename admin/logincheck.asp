
<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>����ʥ���ܽ���Ʒ���޹�˾</title>
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
   ErrMsg=ErrMsg&"�û������ԣ�\n\n"
End if

if len(apass)>20 or len(apass)<3  then
   FoundErr=True
   ErrMsg=ErrMsg&"�û����벻�ԣ�\n\n"
End if
if Safecode="" then
	FoundErr=True
	ErrMsg=ErrMsg & "��֤�벻��Ϊ�գ�\n\n"
end if
if Session("Admin_GetCode")="" then
	FoundErr=True
	ErrMsg=ErrMsg & "���¼ʱ������������·��ص�¼ҳ����е�¼��\n\n"
end if
if Safecode<>CStr(Session("Admin_GetCode")) then
	FoundErr=True
	ErrMsg=ErrMsg & "�������ȷ�����ϵͳ�����Ĳ�һ�£����������롣\n\n"
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
          ErrMsg="�û��������������"
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
