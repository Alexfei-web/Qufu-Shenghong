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
<%
select case Request("action")
	Case "del"
	     Call del()
	Case "show"
	     Call show()	 
	case else
	    Call main()
	End select
        Call CloseDatabase()                                   
%>



<% sub main() %>                                      <!---------------�����б���------------>
<script language="javascript">
    function CheckAll(form)  {
      for (var i=0;i<form.elements.length;i++)    {
         var e = form.elements[i];
         if (e.type.toLowerCase()=="checkbox"&&e.name != 'chkall')
    	    e.checked = form.chkall.checked; 
      }
    }
</script>
<TABLE border=0 width="80%" align=center Class=TableBorder>
  <TR> 
    <TD class=fl2 colspan="2"><b>˵����</b><br>
      (1)һ��ɾ����¼�����޷��ָ��������ؿ��ǣ�</TD>
  </TR>
 
</TABLE>
<br>

<table width="80%"  border="0" align="center" cellpadding="2" cellspacing="1" Class=TableBorder>
<form action="admin_comment.asp?action=del" name=webmate method="post">
  <tr height="25"> 
    <th width="44%">����</th>
    <th width="16%">��ϵ��</th>
	<th width="20%">��ϵ�绰</th>
    <th width="11%">���ʱ��</th>
	
	<th width="9%">ѡ��</th>
  </tr>
<%
dim page
dim rs,sql

Set rs= Server.CreateObject("ADODB.Recordset")
sql="select * from comment "
rs.open sql,conn,1,1
if rs.eof then
response.write "<tr><td class=fl1 colspan=""7"">��û���κ����ԣ�</td></tr>" 
End if

 if request("page")<>"" then  
    Page = cint(Request.QueryString("page"))                                 'pageֵΪ����ֵ 
 else
    Page=1
 end if 
const PageSize = 20                                                          'ÿҳ��ʾ��¼��
rs.PageSize=PageSize

 
    if Page > rs.PageCount then                                              '������յ�ҳ��������ҳ��
        Page = rs.PageCount                                                  '���õ�ǰ��ʾҳ�������ҳ        
    end if 
	if Page <= 0 then                                                        '���pageС�ڵ���0
        Page = 1                                                             '����PAGE���ڵ�һҳ
    end if
        rs.AbsolutePage = Page                                               '���������,��ʾ��ǰҳ���ڽ��յ�ҳ�� 
    


dim i,flcss
i = 0
do while not rs.eof
if i mod 2=0 then
flcss="fl1"
else
flcss="fl2"
end if
%>
  <tr> 
    <td class=<%=flcss%>>
	<a href="admin_comment.asp?action=show&id=<%=rs("id")%>"><%=rs("title")%></a></td>
    <td class=<%=flcss%> align="center"><%=rs("receiver")%></td>
    
	<td class=<%=flcss%> align="center"><%=rs("phone")%></td>
	<td class=<%=flcss%> align="center"><%=rs("time")%></td>
	<td class=<%=flcss%> align="center"><input name="id" type="checkbox" id="id" value="<%=rs("id")%>"></td>
  </tr>
  <%
    i=i+1
	if i>=PageSize then exit do
	rs.movenext
	loop
  %>
<tr> 
    <td colspan="8" Class=fl2> 
        <div align="right"> 
            <table width=""100%""><tr><td align=right>
            <tr><td>
            <a href="admin_comment.asp?page=1">��ҳ</a> | 
            <a href="admin_comment.asp?page=<%= Page-1 %>">��һҳ</a> | 
            <a href="admin_comment.asp?page=<%= Page+1 %>">��һҳ</a> |
            <a href="admin_comment.asp?page=<%= rs.pagecount %>">βҳ</a>
              
            &nbsp;��<%= rs.RecordCount %>��
            
            &nbsp;&nbsp;<input type="checkbox" name="chkall" onClick="CheckAll(this.form)">ȫѡ/ȡ��&nbsp;
            &nbsp;<INPUT TYPE="submit" value="ɾ��" name="action" style="width:40pt;height:22" class="button">&nbsp;&nbsp;</td></tr></table>
        </div>
    </td>
</tr>
<%
    rs.close   
    set rs=nothing
%>
</form>
</table>
<% end sub %>


<% sub del()                                 '_______________ɾ����¼��--------------------
Dim idlist,idArr,uid
dim rs,sql,SucDelNum,i
dim arrUploadFiles,intTemp
if Request("id")="" Then
		ErrMsg=ErrMsg+"<li>����ʧ�ܣ���ѡ��������ݽ��в�����"
		Call Error(ErrMsg)
		Exit Sub
End if

idlist=Request("id")
idArr=Split(idlist,",")
SucDelNum=0
for i = 0 to ubound(idarr)
	uid=Clng(idarr(i))
	set rs=server.CreateObject("adodb.recordset")
	sql="delete from comment where id="& uid
	rs.open sql,conn,1,3
	
	SucDelNum=SucDelNum+1
next
call adminsuc("�ɹ�ɾ��&nbsp;<font color=brown>"&SucDelNum&"</font>&nbsp;�����ԣ�")
end sub %>


<% 
sub show()
dim id
id=trim(request.QueryString("id"))
if id="" or not ISNumeric(id) then
Errmsg=Errmsg+"<br>"+"<li>����ʧ�ܣ�������ʧ��"
Call Error(ErrMsg)
exit sub
end if
id=clng(id)
set rs =server.CreateObject("adodb.recordset")
sql="select * from comment where id="&id
rs.open sql,conn,1,1

if rs.eof then
set rs=nothing
Errmsg=Errmsg+"<br>"+"<li>����ʧ�ܣ��Ҳ��������ԣ�"
Call Error(ErrMsg)
exit sub
end if
%>
<form action="" name="myform" method=post>
  <table cellpadding="2" cellspacing="1" border="0" width="80%" class=TableBorder align=center>
    <tr> 
      <th colspan=2 height=23>�ο�����</th>
    </tr>
    <tr>
      <td colspan="2" Class=fl1 align="center"><table width="90%" height="21" border="0" cellpadding="3" cellspacing="0">
       
		<tr>
          <td>&nbsp;&nbsp;&nbsp;&nbsp;���⣺<input name="title" type="text" id="title" value="<%= rs("title") %>" size="50" maxlength="80">
          </td>
        </tr>
		<tr>
          <td>&nbsp;&nbsp;��ϵ�ˣ�<input name="receiver" type="text" id="receiver" value="<%= rs("receiver") %>" size="35" maxlength="30">
         </td>
        </tr> 
		<tr>
          <td>��ϵ�绰��<input name="phone" type="text" id="phone" value="<%= rs("phone") %>" size="35" maxlength="50">
         </td>
        </tr> 
		<tr>
          <td>&nbsp;&nbsp;&nbsp;&nbsp;���棺<input name="fax" type="text" id="fax" value="<%= rs("fax") %>" size="35" maxlength="50">
         </td>
        </tr> 
		<tr>
          <td>&nbsp;&nbsp;&nbsp;&nbsp;���䣺<input name="email" type="text" id="email" value="<%= rs("email") %>" size="50" maxlength="80">
         </td>
        </tr> 
       
        <tr>
          <td align="center"><b>������ϸ��Ϣ</b></td>
        </tr>
        <tr>
          <td>
		  <textarea name="content" id="content" style="width:100%;height:550px;" value=""><%= rs("content") %></textarea>
		  </td>
        </tr>
      </table>
      </td>
    </tr>
    
  </table>
</form> 

<% end sub %>
</BODY></HTML>