<!--#include file="admin_checkuser.asp"-->
<!--#include file="checkpost.asp"-->
<script type="text/javascript" src="kindeditor/kindeditor.js"></script> 
<script type="text/javascript" src="kindeditor/lang/zh_CN.js"></script> 
<script type="text/javascript">

		KindEditor.ready(function(K) {
			var editor1 = K.create('#editor_id',{
			 width : "650px",
             height : "500px", 
             filterMode : false, 
             resizeMode :1 ,
			   
				cssPath : 'KindEditor/plugins/code/prettify.css',
				uploadJson : 'KindEditor/asp/upload_json.asp',
				fileManagerJson : 'KindEditor/asp/file_manager_json.asp',
				allowFileManager : true,
				afterCreate : function() {
					var self = this;
					K.ctrl(document, 13, function() {
						self.sync();
						K('form[name=myform]')[0].submit();
					});
					K.ctrl(self.edit.doc, 13, function() {
						self.sync();
						K('form[name=myform]')[0].submit();
					});
				}
			});
			
		});
	</script>
<HTML><HEAD><title>����ҳ��</title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<!--#include file="css/admin_main.css"-->
<style type="text/css">textarea{ word-break:break-all; resizeMode:0;}</style>
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
<script language = "JavaScript">
    function checkadd(){
        if (document.myform.title.value == '') {
            alert('���ű��ⲻ��Ϊ��');
		    document.myform.title.focus();
		    return false
		}		
	}
</script>

<%

select case Request("action")
	case "Save"
		Call Save()
	case "add"
		call add()
	Case "edit"
	     Call Edit()
	Case "EditSave"
	     Call EditSave()
	Case "passed"
	      Call passed()
	Case "del"
	     Call del()
	case else
	    Call main()
	End select


Call CloseDatabase()
Sub add()                                            '����������������������������������š�������������������������������


%>
<form action="admin_news.asp?action=Save" name="myform" method=post onSubmit="return checkadd();">
  <table cellpadding="2" cellspacing="1" border="0" width="80%" class=TableBorder align=center>
    <tr> 
      <th colspan=2 height=23>�������</th>
    </tr>
    <tr>
      <td colspan="2" Class=fl1 align="center"><table width="90%" height="21" border="0" cellpadding="3" cellspacing="0">
        <tr>
          <td>&nbsp;���ű��⣺<input name="title" type="text" id="title" value="" size="60" maxlength="150">&nbsp;* </td>
        </tr>
       
        <tr>
          <td>&nbsp;������Դ��<input name="author" type="text" id="author" value="" size="40" maxlength="30"></td>
        </tr>
        <tr>
          <td>&nbsp;���Źؼ��֣�<input name="key" type="text" id="key" size="45" maxlength="500">������ã��ָ�</td>
        </tr>
        <tr>
          <td align="center"><textarea name="content" id="editor_id" style="width:70%;height:550px;visibility:hidden;"></textarea>
		 </td>
        </tr>
       
      </table>
      </td>
    </tr>
    
    <tr align="center"> 
      <td colspan="2" Class=fl2> <input type="submit" name="submit" value="�� ��" class=button> 
        &nbsp;&nbsp; <input type=button value="�� ��" onClick="location.href='admin_news.asp'" class=button>      </td>
    </tr>
  </table>
</form>
<%end sub%>

<% 
    Sub Save()                                    '������ź�   ����
    dim rs,sql,founderr
    founderr=false
    dim Title,Content,Author,key
    
    Title=trim(request.form("title"))
    Author=trim(request.form("author"))
    Content=trim(request.form("content"))
    key=trim(request.form("key"))
    
    if Title="" then
    		founderr=true
    		ErrMsg=ErrMsg & "<br><li>���ű��ⲻ��Ϊ��</li>"
    end if
    if Content="" then
    		founderr=true
    		ErrMsg=ErrMsg & "<br><li>�������ݲ���Ϊ��</li>"
    end if
    
    
    if founderr=true then
    Call Error(ErrMsg)
    exit sub
    end if
    
    set rs=server.createobject("adodb.recordset")
    sql="select * from news"
    rs.open sql,conn,1,3
    rs.addnew
    rs("title")=Title
    rs("author")=Author
    rs("time")=year(now) & "-" & month(now)  &"- " & day(now)
    rs("content")=Content
    rs("hits")=0
    rs("key")=key
    
    rs.update
    rs.close
    set rs=nothing
    
    call adminsuc("������ӳɹ���")
    End Sub
%>

<% 
    Sub EditSave()                                 '����༭��������
    dim rs,sql,founderr,id
    dim Title,Content,Author
    
    Title=trim(request.form("title"))
    Author=trim(request.form("author"))
    
    Content=trim(request.form("content"))
    
    key=trim(request.form("key"))
    
    id =request.form("id")
    
    id=clng(id)
    if Title="" then
    		founderr=true
    		ErrMsg=ErrMsg & "<br><li>���ű��ⲻ��Ϊ��</li>"
    end if
    if Content="" then
    		founderr=true
    		ErrMsg=ErrMsg & "<br><li>�������ݲ���Ϊ��</li>"
    end if
    
    
    if founderr=true then
    Call Error(ErrMsg)
    exit sub
    end if
    
    set rs=server.createobject("adodb.recordset")
    
    sql="select * from news where id="& id
    rs.open sql,conn,1,3
    if rs.eof then
    rs.close
    set rs=nothing
    ErrMsg=ErrMsg & "<br><li>�Ҳ�����Ҫ�޸ĵļ�¼</li>"
    Call Error(ErrMsg)
    exit sub
    end if
    rs("title")=Title
    rs("author")=Author
    
    rs("content")=Content
    rs("key")=key
    rs.update
    rs.close
    set rs=nothing
    
    call adminsuc2("�����޸ĳɹ���","admin_news.asp")
    End Sub 
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
  <TR> 
    <TD class=fl1>

<a href="admin_news.asp?action=add"><font color=blue>&nbsp;&nbsp; �������</font></a></TD>
	
  </TR>
</TABLE>
<br>

<table width="80%"  border="0" align="center" cellpadding="2" cellspacing="1" Class=TableBorder>
<form action="admin_news.asp?action=del" name=webmate method="post">
  <tr height="25"> 
    <th width="44%">����</th>
    <th width="20%">��Դ</th>
	<th width="16%">���ʱ��</th>
    <th width="11%">�������</th>
	
	<th width="9%">ѡ��</th>
  </tr>
<%
    dim page
    dim rs,sql
    
    Set rs= Server.CreateObject("ADODB.Recordset")
    sql="select * from news "
    rs.open sql,conn,1,1
    if rs.eof then
    response.write "<tr><td class=fl1 colspan=""7"">��û���κ����ţ�</td></tr>" 
    End if
    
    if request("page")<>"" then  
        Page = cint(Request.QueryString("page"))                               'pageֵΪ����ֵ 
    else
        Page=1
    end if 
    const PageSize = 20                                                        'ÿҳ��ʾ��¼��
    rs.PageSize=PageSize 
        if Page > rs.PageCount then                                            '������յ�ҳ��������ҳ��
            Page = rs.PageCount                                                '���õ�ǰ��ʾҳ�������ҳ        
        end if 
    	if Page <= 0 then                                                      '���pageС�ڵ���0
            Page = 1                                                           '����PAGE���ڵ�һҳ
        end if
            rs.AbsolutePage = Page                                             '���������,��ʾ��ǰҳ���ڽ��յ�ҳ�� 
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
	<a href="?action=edit&id=<%=rs("id")%>"><%=rs("title")%></a></td>
    <td class=<%=flcss%> align="center"><%=rs("author")%></td>
    <td class=<%=flcss%> align="center"><%=rs("time")%></td>
	<td class=<%=flcss%> align="center"><%= rs("hits") %></td>
	
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
<a href="admin_news.asp?page=1">��ҳ</a> | 
<a href="admin_news.asp?page=<%= Page-1 %>">��һҳ</a> | 
<a href="admin_news.asp?page=<%= Page+1 %>">��һҳ</a> |
<a href="admin_news.asp?page=<%= rs.pagecount %>">βҳ</a>
  
&nbsp;��<%= rs.RecordCount %>��

&nbsp;&nbsp;<input type="checkbox" name="chkall" onClick="CheckAll(this.form)">ȫѡ/ȡ��&nbsp;
&nbsp;<INPUT TYPE="submit" value="ɾ��" name="action" style="width:40pt;height:22" class="button">&nbsp;&nbsp;</td></tr></table>
</div></td>
  </tr>
<%
    rs.close   
    set rs=nothing
%>
</form>
</table>
<% end sub %>


<% 
    sub del()                                 'ɾ����¼��
    Dim idlist,idArr,uid
    dim rs,sql,SucDelNum,i
    dim arrUploadFiles,intTemp
    if Request("id")="" Then
		ErrMsg=ErrMsg+"<li>����ʧ�ܣ���ѡ��������Ž��в�����"
		Call Error(ErrMsg)
		Exit Sub
    End if

    idlist=Request("id")
    idArr=Split(idlist,",")
    SucDelNum=0
    for i = 0 to ubound(idarr)
	uid=Clng(idarr(i))
	set rs=server.CreateObject("adodb.recordset")
	sql="delete from news where id="& uid
	rs.open sql,conn,1,3
	
	SucDelNum=SucDelNum+1
    next
    call adminsuc("�ɹ�ɾ��&nbsp;<font color=brown>"&SucDelNum&"</font>&nbsp;�����ţ�")
    end sub 
%>



<% 
    Sub Edit()                              '�༭������
    dim id,sql,rs
    id=trim(request("id"))
    if id="" or not ISNumeric(id) then
    Errmsg=Errmsg+"<br>"+"<li>����ʧ�ܣ�������ʧ��"
    Call Error(ErrMsg)
    exit sub
    end if
    id=clng(id)
    set rs =server.CreateObject("adodb.recordset")
    sql="select * from news where id="&id
    rs.open sql,conn,1,1
    
    if rs.eof then
    set rs=nothing
    Errmsg=Errmsg+"<br>"+"<li>����ʧ�ܣ��Ҳ��������ţ�"
    Call Error(ErrMsg)
    exit sub
    end if
%>
<form action="admin_news.asp?action=EditSave" name="myform" method=post onSubmit="return checkadd();">
  <table cellpadding="2" cellspacing="1" border="0" width="80%" class=TableBorder align=center>
    <tr> 
      <th colspan=2 height=23>�����޸�</th>
    </tr>
    <tr>
      <td colspan="2" Class=fl1 align="center"><table width="90%" height="21" border="0" cellpadding="3" cellspacing="0">
        <tr>
          <td>&nbsp;���ű��⣺<input name="title" type="text" id="Title" value="<%=rs("title")%>" size="60" maxlength="150">&nbsp;* </td>
        </tr>
       
        <tr>
          <td>&nbsp;������Դ��<input name="author" type="text" id="Author" value="<%=rs("author")%>" size="20" maxlength="30">
		  <input name="id" type="hidden" id="id" value="<%=rs("id")%>"></td></tr>
            
        
        <tr>
          <td>&nbsp;���Źؼ��֣�<input name="key" type="text" id="key" size="40" maxlength="500" value="<%=rs("key")%>"> 
          ������� , �ָ�</td>
        </tr>
        <tr>
          <td align="center"><textarea name="content" id="editor_id" style="width:100%;height:550px;visibility:hidden;"><%=rs("content")%></textarea>
		  </td>
        </tr>
    
      </table>
      </td>
    </tr>
    
    <tr align="center"> 
      <td colspan="2" Class=fl2> <input type="submit" name="Submit" value="�� ��" class=button> 
        &nbsp;&nbsp; <input type=button value="�� ��" onClick="location.href='admin_news.asp'" class=button>      </td>
    </tr>
  </table>
</form>
<%
set rs=nothing
end sub
%>

</BODY></HTML>