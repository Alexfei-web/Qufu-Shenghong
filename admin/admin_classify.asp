<!--#include file="admin_checkuser.asp"-->
<html>
<head>
    <title>����ҳ��</title>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312">
    <!--#include file="css/admin_main.css"-->
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
Select case request("action")
  case "add"
    call  add()
  case "addsave"
    call  save()
  case "edit"
    Call edit()
  case "editsave"
    Call editsave()
  case "del"
    Call del()
  case else
    call main()
End Select


Sub editsave()
dim classname,classid
dim sql,rs

classid=Clng(request.form("classid"))     '�ѱ��ʽת��Ϊ�����Σ�Long������
classname=trim(request.Form("classname"))

set rs=server.createobject("adodb.recordset")
sql="select * from classify Where id="&Clng(Request.Form("id"))
rs.open sql,conn,1,3
if not rs.eof then
rs("classname")=classname
rs("classid")=Clng(trim(request.form("classid")))
rs.update
end if
rs.close
set rs=nothing
response.redirect "admin_classify.asp"
End Sub



 
sub save()            '-------------��ӱ���-------------
dim classname,classid
dim sql,rs

classname=trim(Request.form("classname"))
if classname="" Then
	ErrMsg=ErrMsg+"<br>"+"<li>����ʧ�ܣ�����д�������ƣ�"
       Call Error(ErrMsg)
       exit sub
End if
   classid=Clng(Request.Form("classid"))
   set rs=server.createobject("adodb.recordset")
   sql="select * from classify where classid="&classid
   rs.open sql,conn,1,3
   if not rs.eof then
     ErrMsg=ErrMsg+"<br>"+"<li>����ʧ�ܣ��Ѵ��ڴ�ClassID��"
       Call Error(ErrMsg)
       exit sub
	End if

set rs=server.createobject("adodb.recordset")
sql="select * from classify"
rs.open sql,conn,1,3
rs.addnew
rs("classid")=classid
rs("classname")=classname
rs.update

rs.close
set rs=nothing
response.redirect "admin_classify.asp"
end sub



sub del()
dim sql,rs,id
id=clng(request("id"))
set rs=server.createobject("adodb.recordset")
sql="select * from classify where id="&id
rs.open sql,conn,1,1
if rs.eof and rs.bof then
	ErrMsg=ErrMsg+"<br>"+"<li>����ʧ�ܣ�û���ҵ�Ҫ���в����Ķ���"
    Call Error(ErrMsg)
	exit sub
end if
set rs=server.createobject("adodb.recordset")
sql="delete from classify where id="&id
rs.open sql,conn,1,3
set rs=nothing

response.redirect "admin_classify.asp"
end sub


sub main()              '--------���б���----------
%>


<%
dim rs,sql
set rs=server.CreateObject("adodb.recordset")
sql="select * from classify "

rs.open sql,conn,1,1 
%>
    <table border="0" width="80%" align="center" class="TableBorder">
        <tr>
            <td colspan="2" class="fl2"><b>˵����</b><br>
                (1)һ��ɾ�������޷��ָ��������ؿ��ǣ�<br>
                (2)<a href="admin_classify.asp?action=add" style="text-decoration: underline; color: #0000FF;">��ӷ���</a>
            </td>
        </tr>
        <br>
    </table>
    <br>
    <table width="80%" border="0" align="center" cellpadding="2" cellspacing="1" class="TableBorder">
        <tr>
            <th colspan="5" height="25">��Ʒ�������</th>
        </tr>
        <%if rs.eof then%>
        <tr>
            <td colspan="5" class="fl1">��û���κη���</td>
        </tr>
        <%else%>
        <tr align="center">
            <td width="300" height="25" class="fl2">�������</td>
            <td width="200" height="25" class="fl2">ClassID</td>
            <td width="200" class="fl2">�޸�</td>
            <td class="fl2">ɾ��</td>
        </tr>
        <%do until rs.eof %>
        <form name="form" method="post" action="">

            <tr>

                <td height="25" align="center" class="fl1">
                    <input name="classname" id="ClassName" value="<%= rs("classname") %>" size="30" maxlength="50"></td>
                <td height="25" align="center" class="fl1">
                    <input name="classid" id="ClassID" value="<%= rs("classid") %>"></td>
                <td align="center" class="fl1"><a href="admin_classify.asp?action=edit&id=<%= rs("id") %>">
                    <img src="images/xinhua10.jpg" width="77" height="21" border="0"></a></td>
                <td align="center" class="fl1"><a href="admin_classify.asp?action=del&id=<%= rs("id") %>">ɾ��</a></td>
            </tr>
        </form>
        <%
            rs.movenext
			loop
			End if
        %>
    </table>
    <%
        rs.close
        set rs=nothing
    %>
    <br>
    <%End Sub%>




    <%
        sub edit()   '-------------�༭��-------
        dim id
        id=request("id")
        set rs=server.CreateObject("adodb.recordset")
        sql="select * from classify where id="&id
        rs.open sql,conn,1,3
    %>
    <script language="javascript">
        function CheckForm(form) {
            if (form.classname.value == "") {
                alert("������������ƣ�");
                form.classname.focus();
                return false;
            }
            if (form.classid.value == "") {
                alert("������ClassID��");
                form.classid.focus();
                return false;
            }
        }
    </script>
    <table border="0" width="80%" align="center" class="TableBorder">
        <tr>
            <td colspan="2" class="fl2"><b>˵����</b><br>
                (1)ClassID�����������������
            </td>
        </tr>

    </table>
    <br>
    <table width="80%" border="0" align="center" cellpadding="2" cellspacing="1" class="TableBorder">
        <tr>
            <th colspan="5" height="25">�޸ķ���</th>
        </tr>
        <form name="form1" method="post" action="admin_classify.asp?action=editsave" onsubmit="return CheckForm(this)">

            <tr>
                <td height="25" align="center" class="fl2" width="300">�������</td>

                <td align="center" class="fl2" width="300">ClassID</td>
                <td align="center" class="fl2">�޸�</td>

            </tr>
            <tr>
                <td height="25" align="center" class="fl1">
                    <input name="id" type="hidden" value="<%= id %>"><input name="classname" type="text" class="form2" id="classname" size="30" value="<%= rs("classname") %>">
                    <font color="#FF0000">*</font></td>
                <td align="center" class="fl1">
                    <input name="classid" type="text" class="form2" id="classif" value="<%= rs("classid") %>" size="15">
                    <font color="#FF0000">&nbsp; *����Ϊ������</font></td>
                <td align="center" class="fl1">
                    <input type="submit" name="Submit" value="" style="background: url(images/xinhua10.jpg); width: 77px; height: 21px;"></td>
            </tr>
        </form>
    </table>
    <% end sub %>


    <% sub add()%>


    <script language="javascript">
        function CheckForm(form) {
            if (form.classname.value == "") {
                alert("������������ƣ�");
                form.classname.focus();
                return false;
            }
            if (form.classid.value == "") {
                alert("������ClassID��");
                form.classid.focus();
                return false;
            }

        }
    </script>
    <table border="0" width="80%" align="center" class="TableBorder">
        <tr>
            <td colspan="2" class="fl2"><b>˵����</b><br>
                (1)һ��ɾ�������޷��ָ��������ؿ��ǣ�<br>
                (2)ClassID�����������������
            </td>
        </tr>


    </table>
    <br>
    <table width="80%" border="0" align="center" cellpadding="2" cellspacing="1" class="TableBorder">
        <tr>
            <th colspan="5" height="25">��ӷ���</th>
        </tr>
        <form name="form1" method="post" action="admin_classify.asp?action=addsave" onsubmit="return CheckForm(this)">

            <tr>
                <td height="25" align="center" class="fl2" width="300">�������</td>

                <td align="center" class="fl2" width="300">ClassID</td>
                <td align="center" class="fl2">���</td>

            </tr>
            <tr>
                <td height="25" align="center" class="fl1">
                    <input name="classname" type="text" class="form2" id="classname" size="30">
                    <font color="#FF0000">*</font></td>
                <td align="center" class="fl1">
                    <input name="classid" type="text" class="form2" id="classif" value="" size="15">
                    <font color="#FF0000">&nbsp; *����Ϊ������</font></td>
                <td align="center" class="fl1">

                    <input type="submit" name="Submit" value="" style="background: url(images/xinhua9.jpg); width: 74px; height: 24px;">
                </td>
            </tr>
        </form>
    </table>
    <% end sub %>
</body>
</html>
