<!--#include file="admin_checkuser.asp"-->
<!--#include file="checkpost.asp"-->
<script type="text/javascript" src="kindeditor/kindeditor.js"></script> 
<script type="text/javascript" src="kindeditor/lang/zh_CN.js"></script> 
<script type="text/javascript">
        //���س�ʼ�������߱༭��
		KindEditor.ready(function(K) {
			var editor1 = K.create('#editor_id', {
			    width : "90%",
                height : "500px", 
                filterMode : false,             //������html����
                resizeMode : 0 ,                //ֻ�ܵ����߶�
				cssPath : 'KindEditor/plugins/code/prettify.css',  //css·��
				uploadJson : 'KindEditor/asp/upload_json.asp',     //�ϴ��ķ���������
				fileManagerJson : 'KindEditor/asp/file_manager_json.asp',  //ָ�����Զ��ͼƬ�ķ������˳���
				allowFileManager : true,                               //��ʾ���������ͼƬ����
				afterCreate : function() {                            //���ñ༭��������ִ�еĻص�������
					var self = this;
					K.ctrl(document, 13, function() {
						self.sync();
						K('form[name=form1]')[0].submit();
					});
					K.ctrl(self.edit.doc, 13, function() {
						self.sync();
						K('form[name=form1]')[0].submit();
					});
				}
			});
		});
</script>

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
select case request("action")                                     '�ж���Ϊ
	case "Save"
		Call save()
	case else
		Call main()
	End select
    Call CloseDatabase()

    Sub main()
    dim rs
    set rs=server.CreateObject("adodb.recordset")
    sql="select * from aboutus"
    rs.open sql,conn,1,1
%> 

<form action="admin_aboutus.asp?action=Save" method=post>
  <table cellpadding="2" cellspacing="1" border="0" width="80%" class=TableBorder align=center>
    <tr> 
      <th colspan=2 height=23>�������ǹ���</th>
    </tr>
    <tr> 
      <td colspan="2" Class=fl1 align="center">
          <table width="90%" height="21" border="0" cellpadding="0" cellspacing="0">
            <tr>
                <td align="center"><textarea name="content" id="editor_id" style="width:90%;height:550px;visibility:hidden;"><%=rs("content")%></textarea></td>
             </tr>
          </table>
      </td>
    </tr>
    
    <tr align="center"> 
      <td colspan="2" Class=fl2> <input type="submit" name="Submit" value="�� ��" class=button> 
        &nbsp;&nbsp; <input type=button value="�� ��" onClick="location.href='admin_aboutus.asp'" class="button">      </td>
    </tr>
    <tr align="center">
      <td height="26" colspan="2" Class=fl2>����ı༭���а�Ť��<img src="images/img.gif" width="20" height="20">���Դӱ����������ϴ�ͼƬ�������С� </td>
    </tr>
  </table>
</form>
<%
    set rs=nothing                                       '��ʽ�����ñ���Ϊ"��"
    end sub
    
    Sub save()
    dim sql,rs,msg
    set rs=server.createobject("adodb.recordset")
    sql="select * from aboutus"
    rs.open sql,conn,1,3                                '�޸�����
    rs("content")=trim(request.form("content"))
    rs.update
    rs.close
    set rs=nothing
    
    call adminsuc("�޸ĳɹ���")
    End Sub
%>
</BODY></HTML>