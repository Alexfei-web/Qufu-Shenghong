<!--#include file="admin_checkuser.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="checkpost.asp"-->
<HTML><HEAD><title>����ҳ��</title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<!--#include file="css/admin_main.css"-->
</HEAD>
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
    end if
%>
<BODY leftMargin=0 topMargin=0 marginwidth="0" marginheight="0">
<%
select case request("action")  '�ж���Ϊ
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
      <B><a name=top>˵��</a></B>��<br>
      (1)ɾ������Ա�Ĳ�����ɾ����̨����Ա�ʺţ����ɻָ���<br>
      (2)Ҫ�޸Ĺ���Ա�ĺ�̨��½�ʺ�������Ӧ���û�����<br>
      </td>
  </tr>
  <tr>
    <td class="fl1"><strong>����������</strong><a href=admin_admin.asp?action=NewAdmin><font color=blue>��ӹ���Ա</font></a></td>
  </tr>
</table>
<table width=80% align=center cellspacing=1 cellpadding=5 Class=TableBorder>
  <tr> 
    <th width="18%">�û���</th>
    <th width="25%">�������</th>
	<th width="9%">��½����</th>
    
    <th width="28%">����½ʱ��</th>
    
    <th width="20%">����</th>
  </tr>
<%
    dim sql,rs,i,css
    set rs=server.CreateObject("adodb.recordset")
    sql="select username,logins,lastlogin from admin"
    rs.open sql,conn,1,1
    
    if rs.eof then 
%>
<tr><td>��û�й���Ա��</td></tr>
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
    <td align=center Class=<%=css%>> ϵͳ����Ա </td>
    <td Class=<%=css%> align=center><%=rs("logins")%></td>
    <td align=center Class=<%=css%>> <%=rs("lastlogin")%></td>
    
    <td align=center Class=<%=css%>><a href="admin_admin.asp?action=Delete&username=<%=rs("username")%>" onClick="return confirm('��ȷ��Ҫɾ����\n\n�뱣��һλϵͳ����Ա��\n�����޷������̨��')">ɾ��</a> 
  </tr>
<%
    rs.movenext '�ƶ������ݱ���һ����¼
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
<tr><th colspan=2>��ӹ���Ա</th></tr>
<tr><td width=45% align=right Class=fl1>�û���&nbsp;</td><td Class=fl1><input name="username" type="text" size=30>&nbsp;</td>
</tr>
<tr><td width=45% align=right Class=fl2>��½����&nbsp;</td><td Class=fl2><input name="userpwd" type="text" size=30></td></tr>
<tr><td colspan=2 align=center Class=fl2><input type="submit" name="Submit" value="&nbsp;��&nbsp;��&nbsp;" class=button>
</td></tr></table>
</form>

<%End Sub%>

<%  
    Sub addadmin() 
    dim sql,rs
	Dim username,userpwd
	username=replace(Request.form("username"),"'","")
	if username="" Then
		ErrMsg=ErrMsg+"<br>"+"<li>����ʧ�ܣ�����д�û�����"
        Call Error(ErrMsg)
        exit sub
	End if
	
	userpwd=trim(Request.form("userpwd"))
	if userpwd="" Then
		ErrMsg=ErrMsg+"<br>"+"<li>����ʧ�ܣ�����д��½���룡"
        Call Error(ErrMsg)
		exit sub
	Else
		userpwd=md5(userpwd)                               '���м���
	End if
    set rs=server.CreateObject("adodb.recordset")
	sql="select * from admin where username='"&username&"'"
	rs.open sql,conn,1,3
	
	if not (rs.eof and rs.bof) Then
		ErrMsg=ErrMsg+"<br>"+"<li>����ʧ�ܣ����û��Ѿ��ǹ���Ա��"
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
    ErrMsg=ErrMsg+"<br>"+"<li>����ʧ�ܣ�������ʧ��"
        Call Error(ErrMsg)
		exit sub
    end if
    set rs=server.CreateObject("adodb.recordset")
	sql="select username from admin where username='"&username&"'"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		ErrMsg=ErrMsg+"<br>"+"<li>����ʧ�ܣ�û���ҵ�Ҫ���в����Ķ���"
    Call Error(ErrMsg)
	else
	    set rs=server.CreateObject("adodb.recordset")
		sql="delete from admin where username='"&username&"'"
		rs.open sql,conn,1,3
		call adminsuc("ɾ���ɹ�������")
		
	end if
	set rs=nothing
    end sub 
%> 
</BODY></HTML>