<!--#include file="admin_checkuser.asp"-->
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


<script LANGUAGE="javascript">
    function Check() {
        if (document.contact.addr.value=="")
        	{
        	  alert("请输入公司地址！")
        	  document.contact.addr.focus()
        	  return false
        	 }
        if (document.contact.phone.value=="")
        	{
        	  alert("请输入公司电话！")
        	  document.contact.phone.focus()
        	  return false
        	 }
        if (document.contact.email.value=="")
        	{
        	  alert("请输入公司邮箱！")
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
      <th height="22" colspan="2" align="center"><strong>联系我们管理</strong></th>
    </tr>
   
    <tr> 
      <td width="40%" height="25" valign="middle" Class=fl1><div align="right"><strong>地址：</strong></div></td>
      <td width="60%" height="25" Class=fl1> <input name="addr" size="40"  maxlength="50" title="您的公司地址" value="<%= rs("addr") %>"> 
        <font color="#FF0000"> *</font></td>
    </tr>
    <tr> 
      <td width="40%" height="25" Class=fl1><div align="right"><strong>电话：</strong></div></td>
      <td height="25" Class=fl1> <input name="phone" size="30"  maxlength="40" type="text"  value="<%= rs("phone") %>" title="您的公司电话"> 
        <font color="#FF0000">*</font></td>
    </tr>
	<tr> 
      <td width="40%" height="25" Class=fl1><div align="right"><strong>传真：</strong></div></td>
      <td height="25" Class=fl1> <input name="fax" size="30"  maxlength="40" type="text"  value="<%= rs("fax") %>" title="您的公司传真"> 
      </td>
    </tr>
	<tr> 
      <td width="40%" height="25" Class=fl1><div align="right"><strong>邮箱：</strong></div></td>
      <td height="25" Class=fl1> <input name="email" size="40"  maxlength="50" type="text"  value="<%= rs("email") %>" title="您的公司邮箱"> 
        <font color="#FF0000">*</font></td>
    </tr>
		<tr> 
      <td width="40%" height="25" Class=fl1><div align="right"><strong>邮编：</strong></div></td>
      <td height="25" Class=fl1> <input name="yb" size="30"  maxlength="40" type="text"  value="<%= rs("yb") %>" title="邮政编码"> 
      </td>
    </tr>
			<tr> 
      <td width="40%" height="25" Class=fl1><div align="right"><strong>服务电话：</strong></div></td>
      <td height="25" Class=fl1> <input name="sj" size="30"  maxlength="40" type="text"  value="<%= rs("sj") %>" title="手机"> 
      </td>
    </tr>

   <tr> 
      <td width="40%" height="25" Class=fl1><div align="right"><strong>备案号：</strong></div></td>
      <td height="25" Class=fl1> <input name="recordno" size="40"  maxlength="1000" type="text"  value="<%= rs("recordno") %>" title="公司备案号"> 
        </td>
    </tr>
    <tr> 
      <td height="40" colspan="2" align="center" Class=fl2>
        <input type="submit" value=" 确 定 " name="cmdOk"> &nbsp; <input type="reset" value=" 重 填 " name="cmdReset">      </td>
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
		ErrMsg=ErrMsg & "<br><li>公司地址不能为空！</li>"
	else
		if phone="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>电话不能为空！</li>"
		end if
	end if
	if email="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>公司邮箱不能为空！</li>"
		
	end if
	 if FoundErr=True then
	 Call Error(ErrMsg)
     exit sub
     end if
	
	
	set rs=Server.CreateObject("Adodb.RecordSet")
	sql="select top 1 * from contact"
	rs.open sql,conn,1,3
	if rs.bof and rs.eof then
		ErrMsg="<br><li>出现未知错误！</li>"
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
	call adminsuc("联系我们修改成功！！")
    end sub
%>
</body>
</html>