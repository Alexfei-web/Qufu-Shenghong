<!--#include file="admin_checkuser.asp"-->
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


<script LANGUAGE="javascript">
    function Check() {
        if (document.contact.addr.value=="")
        	{
        	  alert("�����빫˾��ַ��")
        	  document.contact.addr.focus()
        	  return false
        	 }
        if (document.contact.phone.value=="")
        	{
        	  alert("�����빫˾�绰��")
        	  document.contact.phone.focus()
        	  return false
        	 }
        if (document.contact.email.value=="")
        	{
        	  alert("�����빫˾���䣡")
        	  document.contact.email.focus()
        	  return false
        	 }
    }
</script>
<div style="height:100px;"></div>
<br>
<%
select case request("action")
     case "save"
	     call save()
     case else
         call main()
     end select
         Call CloseDatabase()

sub main()
set rs=server.CreateObject("adodb.recordset")
sql="select * from contact"
rs.open sql,conn,1,3
if rs.eof then
rs.addnew
rs("addr")=""
rs("phone")=""
rs("fax")=""
rs("email")=""
rs("recordno")=""
rs("yb")=""
rs("sj")=""
rs.update
rs.close
set rs=nothing
end if
%>
<% 
    set rs=server.CreateObject("adodb.recordset")
    sql="select top 1 * from contact"
    rs.open sql,conn,1,1 
%>
<form method="post" name="contact" onSubmit="return Check()" action="admin_contact.asp?action=save">
  <table border="0" cellpadding="2" cellspacing="1" align="center" width="100%" class=TableBorder>
    <tr> 
      <th height="22" colspan="2" align="center"><strong>��ϵ���ǹ���</strong></th>
    </tr>
   
    <tr> 
      <td width="40%" height="25" valign="middle" Class=fl1><div align="right"><strong>��ַ��</strong></div></td>
      <td width="60%" height="25" Class=fl1> <input name="addr" size="40"  maxlength="50" title="���Ĺ�˾��ַ" value="<%= rs("addr") %>"> 
        <font color="#FF0000"> *</font></td>
    </tr>
    <tr> 
      <td width="40%" height="25" Class=fl1><div align="right"><strong>�绰��</strong></div></td>
      <td height="25" Class=fl1> <input name="phone" size="30"  maxlength="40" type="text"  value="<%= rs("phone") %>" title="���Ĺ�˾�绰"> 
        <font color="#FF0000">*</font></td>
    </tr>
	<tr> 
      <td width="40%" height="25" Class=fl1><div align="right"><strong>���棺</strong></div></td>
      <td height="25" Class=fl1> <input name="fax" size="30"  maxlength="40" type="text"  value="<%= rs("fax") %>" title="���Ĺ�˾����"> 
      </td>
    </tr>
	<tr> 
      <td width="40%" height="25" Class=fl1><div align="right"><strong>���䣺</strong></div></td>
      <td height="25" Class=fl1> <input name="email" size="40"  maxlength="50" type="text"  value="<%= rs("email") %>" title="���Ĺ�˾����"> 
        <font color="#FF0000">*</font></td>
    </tr>
		<tr> 
      <td width="40%" height="25" Class=fl1><div align="right"><strong>�ʱࣺ</strong></div></td>
      <td height="25" Class=fl1> <input name="yb" size="30"  maxlength="40" type="text"  value="<%= rs("yb") %>" title="��������"> 
      </td>
    </tr>
			<tr> 
      <td width="40%" height="25" Class=fl1><div align="right"><strong>����绰��</strong></div></td>
      <td height="25" Class=fl1> <input name="sj" size="30"  maxlength="40" type="text"  value="<%= rs("sj") %>" title="�ֻ�"> 
      </td>
    </tr>

   <tr> 
      <td width="40%" height="25" Class=fl1><div align="right"><strong>�����ţ�</strong></div></td>
      <td height="25" Class=fl1> <input name="recordno" size="40"  maxlength="1000" type="text"  value="<%= rs("recordno") %>" title="��˾������"> 
        </td>
    </tr>
    <tr> 
      <td height="40" colspan="2" align="center" Class=fl2>
        <input type="submit" value=" ȷ �� " name="cmdOk"> &nbsp; <input type="reset" value=" �� �� " name="cmdReset">      </td>
    </tr>
  </table>
</form>
<%
end sub
%>


<%
sub save()
	
	dim addr,phone,fax,email,recordno,yb,sj
	addr=trim(request("addr"))
	phone=trim(request("phone"))
	fax=trim(request("fax"))
	email=trim(request("email"))
	recordno=trim(request("recordno"))
	yb=trim(request("yb"))
	sj=trim(request("sj"))
	
	if addr="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��˾��ַ����Ϊ�գ�</li>"
	else
		if phone="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>�绰����Ϊ�գ�</li>"
		end if
	end if
	if email="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��˾���䲻��Ϊ�գ�</li>"
		
	end if
	 if FoundErr=True then
	 Call Error(ErrMsg)
     exit sub
     end if
	
	
	set rs=Server.CreateObject("Adodb.RecordSet")
	sql="select top 1 * from contact"
	rs.open sql,conn,1,3
	if rs.bof and rs.eof then
		ErrMsg="<br><li>����δ֪����</li>"
		call Error(ErrMsg)
	else
		
       rs("addr")=addr
       rs("phone")=phone
       rs("fax")=fax
       rs("email")=email
       rs("recordno")=recordno
	   rs("yb")=yb
	   rs("sj")=sj
       rs.update
		
	   rs.close
	   set rs=nothing

	end if
	call adminsuc("��ϵ�����޸ĳɹ�����")
    end sub
%>
</body>
</html>