<!--#include file="admin_checkuser.asp"-->
<!--#include file="checkpost.asp"-->
<script type="text/javascript" src="kindeditor/kindeditor.js"></script> 
<script type="text/javascript" src="kindeditor/lang/zh_CN.js"></script> 
<script type="text/javascript">
        //加载初始化的在线编辑器
		KindEditor.ready(function(K) {
			var editor1 = K.create('#editor_id', {
			    width : "90%",
                height : "500px", 
                filterMode : false,             //不过滤html代码
                resizeMode : 0 ,                //只能调整高度
				cssPath : 'KindEditor/plugins/code/prettify.css',  //css路径
				uploadJson : 'KindEditor/asp/upload_json.asp',     //上传的服务器程序
				fileManagerJson : 'KindEditor/asp/file_manager_json.asp',  //指定浏览远程图片的服务器端程序。
				allowFileManager : true,                               //显示浏览服务器图片功能
				afterCreate : function() {                            //设置编辑器创建后执行的回调函数。
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
<%
select case request("action")                                     '判断行为
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
      <th colspan=2 height=23>关于我们管理</th>
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
      <td colspan="2" Class=fl2> <input type="submit" name="Submit" value="更 新" class=button> 
        &nbsp;&nbsp; <input type=button value="返 回" onClick="location.href='admin_aboutus.asp'" class="button">      </td>
    </tr>
    <tr align="center">
      <td height="26" colspan="2" Class=fl2>上面的编辑器中按扭：<img src="images/img.gif" width="20" height="20">可以从本机电脑中上传图片到内容中。 </td>
    </tr>
  </table>
</form>
<%
    set rs=nothing                                       '显式声明该变量为"无"
    end sub
    
    Sub save()
    dim sql,rs,msg
    set rs=server.createobject("adodb.recordset")
    sql="select * from aboutus"
    rs.open sql,conn,1,3                                '修改数据
    rs("content")=trim(request.form("content"))
    rs.update
    rs.close
    set rs=nothing
    
    call adminsuc("修改成功！")
    End Sub
%>
</BODY></HTML>