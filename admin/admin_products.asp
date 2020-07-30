<!--#include file="admin_checkuser.asp"-->
<!--#include file="checkpost.asp"-->
<script type="text/javascript" src="kindeditor/kindeditor.js"></script> 
<script type="text/javascript" src="kindeditor/lang/zh_CN.js"></script> 
<script type="text/javascript">



    KindEditor.ready(function (K) {
        var editor1 = K.create('#editor_id', {
            width: "650px",
            height: "500px",
            filterMode: false,
            resizeMode: 0,
            cssPath: 'KindEditor/plugins/code/prettify.css',
            uploadJson: 'KindEditor/asp/upload_json.asp',
            fileManagerJson: 'KindEditor/asp/file_manager_json.asp',
            allowFileManager: true,
            afterCreate: function () {
                var self = this;
                K.ctrl(document, 13, function () {
                    self.sync();
                    K('form[name=myform]')[0].submit();
                });
                K.ctrl(self.edit.doc, 13, function () {
                    self.sync();
                    K('form[name=myform]')[0].submit();
                });
            }
        });
        
    });
	</script>
<HTML><HEAD><title>管理页面</title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<!--#include file="css/admin_main.css"-->
</HEAD>

<script language = "JavaScript">
    function checkadd() {
        if (document.myform.ProName.value == '') {
            alert('产品名称不能为空');
            document.myform.ProName.focus();
            return false
        }
        if (document.myform.ClassId.value == '' || document.myform.ClassId.value <= 0) {
            alert('请选择产品类型');
            document.myform.ClassId.focus();
            return false
        }
    }
</script>
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

select case Request("action")
	case "Save"
		Call Save()
	case "add"
		call add()
	Case "edit"
	     Call Edit()
	Case "EditSave"
	     Call EditSave()
	
	Case "del"
	     Call del()
	case else
	    Call main()
	End select

Call CloseDatabase()

Sub add()                                '产品添加项

%>
<form action="admin_products.asp?action=Save" name="myform" method=post onSubmit="return checkadd();">
  <table cellpadding="2" cellspacing="1" border="0" width="80%" class=TableBorder align=center>
    <tr> 
      <th colspan=2 height=23>产品添加</th>
    </tr>
    <tr>
      <td colspan="2" Class=fl1 align="center"><table width="90%" height="21" border="0" cellpadding="3" cellspacing="0">
       
		<tr>
          <td>&nbsp;产品名称：<input name="ProName" type="text" id="ProName" value="" size="35" maxlength="50">
          &nbsp;* </td>
        </tr>
		<tr>
          <td>&nbsp;产品型号：<input name="ProType" type="text" id="ProType" value="" size="35" maxlength="30">
         </td>
        </tr> 
		<tr>
          <td>&nbsp;产品类型：<select name="ClassId" size="1" id="ClassId">
             <option value="0" selected="selected">请选择</option>
			 <% 
                 set rs2=server.CreateObject("adodb.recordset")
				 sql="select * from classify"
				 rs2.open sql,conn,1,1
				 do while not rs2.eof
			 %>
             <option value="<%= rs2("classid") %>"><%= rs2("classname") %></option>
			 
			<%  
                rs2.movenext
				loop
				rs2.close
            %> 
			 
        </select>
            &nbsp;*	</td>
        </tr>
        <tr>
          <td>&nbsp;产品添加：<input name="AddUser" type="text" id="AddUser" value="" size="20" maxlength="30"></td>
        </tr>
        <tr>
          <td align="center"><b>产品详细信息</b></td>
        </tr>
        <tr>
          <td align="center"><!--<INPUT type="hidden" name="ProContent" id="Content" style="display:none"> -->
		  <textarea name="ProContent" id="editor_id" style="width:100%;height:550px;visibility:hidden;"></textarea>
		  </td>
        </tr>
       
        <tr>
          <td>产品图片：<input name="DefaultPicUrl" type="text" id="DefaultPicUrl" size="50" maxlength="200">
              用于产品预览显示 <br><iframe id=pic_file frameborder=0 src="uploadfiles.asp" width="100%" height="50" scrolling=no></iframe></td>
        </tr>
	
      </table>
      </td>
    </tr>
    
    <tr align="center"> 
      <td colspan="2" Class=fl2> <input type="submit" name="Submit" value="发 布" class=button> 
        &nbsp;&nbsp; <input type=button value="返 回" onClick="location.href = 'admin_products.asp'" class=button>      </td>
    </tr>
  </table>
</form>
<%
end sub

Sub Save()                          '产品添加后  保存项
dim rs,sql,founderr
dim ProType,ProName,ClassId
dim AddUser,content,ProPrice
dim DefaultPicUrl
ProID=trim(request.form("ProType"))
ProName=trim(request.form("ProName"))
ClassId=trim(request.form("ClassId"))
AddUser=trim(request.form("AddUser"))
ProContent=trim(request.form("ProContent"))

DefaultPicUrl=trim(request.form("DefaultPicUrl"))
if DefaultPicUrl="" then DefaultPicUrl="upload/nopic.gif"

if ProName="" then
		founderr=true
		ErrMsg=ErrMsg & "<br><li>产品名称不能为空</li>"
end if
if ClassId="" or ClassId="0" then
		founderr=true
		ErrMsg=ErrMsg & "<br><li>请选择产品类型</li>"
else
ClassId=clng(ClassId)
end if

if ProContent="" then
		founderr=true
		ErrMsg=ErrMsg & "<br><li>产品介绍不能为空</li>"
end if


if founderr=true then
Call Error(ErrMsg)
exit sub
end if


set rs=server.createobject("adodb.recordset")
sql="select * from products"
rs.open sql,conn,1,3

rs.addnew
rs("ProType")=ProType
rs("ProName")=ProName
rs("ClassId")=ClassId
rs("AddUser")=AddUser
rs("ProContent")=ProContent
rs("Time")=year(now) & "-" & month(now)  &"- " & day(now)
rs("Hits")=0
rs("ProPic")=DefaultPicUrl

rs.update
rs.close
set rs=nothing

call adminsuc("产品添加成功！")

End Sub 


Sub EditSave()                                  '编辑保存项
dim rs,sql,founderr,id
dim ProType,ProName,ClassId
dim AddUser,content,DefaultPicUrl

ProType=trim(request.form("Protype"))
ProName=trim(request.form("ProName"))
ClassId=trim(request.form("ClassId"))
AddUser=trim(request.form("AddUser"))
ProContent=trim(request.form("ProContent"))
DefaultPicUrl=trim(request.form("DefaultPicUrl"))
id=trim(request.form("pid"))

%>
  
<%
if id="" or not ISNumeric(id) then
Errmsg=Errmsg+"<br>"+"<li>操作失败，参数丢失！"
Call Error(ErrMsg)
exit sub
end if
id=clng(id)
if DefaultPicUrl="" then DefaultPicUrl="upload/nopic.gif"

if ProName="" then
	founderr=true
	ErrMsg=ErrMsg & "<br><li>产品名称不能为空</li>"
end if
if ClassId="" or ClassId="0" then
	founderr=true
	ErrMsg=ErrMsg & "<br><li>请选择产品类型</li>"
else
ClassId=clng(ClassId)
end if

if ProContent="" then
	founderr=true
	ErrMsg=ErrMsg & "<br><li>产品介绍不能为空</li>"
end if

if founderr=true then
Call Error(ErrMsg)
exit sub
end if

set rs=server.createobject("adodb.recordset")
sql="select * from products where id="&id&""
rs.open sql,conn,1,3
if rs.eof then
rs.close
set rs=nothing
ErrMsg=ErrMsg & "<br><li>找不到需要修改的记录</li>"
Call Error(ErrMsg)
exit sub
end if
rs("ProType")=ProType
rs("ProName")=ProName
rs("ClassId")=ClassId
rs("AddUser")=AddUser
rs("ProContent")=ProContent
rs("ProPic")=DefaultPicUrl
rs("Time")=year(now) & "-" & month(now)  &"- " & day(now)
rs.update
rs.close
set rs=nothing

call adminsuc2("产品修改成功！","admin_products.asp")
End Sub

%>

<%


sub main()                           '-----------------主列表项----------------
dim ClassId
ClassId=trim(request("ClassId"))
if ClassId="" or ClassId="0" then
ClassId=0
else
ClassId=clng(ClassId)
end if
%>
<script language="javascript">
    function ToChange() {
        var obj = document.getElementById("pro_check");
        var index = obj.selectedIndex;
        obj.options[index].selected = true;
    }

    function CheckAll(form) {
        for (var i = 0; i < form.elements.length; i++) {
            var e = form.elements[i];
            if (e.type.toLowerCase() == "checkbox" && e.name != 'chkall')
                e.checked = form.chkall.checked;
        }
    }
</script>
<form action="admin_products.asp?action=del" name=webmate method="post">
<TABLE border=0 width="80%" align=center Class=TableBorder>
  <TR> 
    <TD class=fl2 colspan="2"><b>说明：</b><br>
      (1)一旦删除记录，就无法恢复，请慎重考虑！</TD>
  </TR>
  <TR> 
  
	<td class=fl1> &nbsp;&nbsp;&nbsp;&nbsp;  <a href="admin_products.asp?action=add"><font color=blue>添加产品</font></a></td>
  </TR>
</TABLE>
<br>

<table width="80%"  border="0" align="center" cellpadding="2" cellspacing="1" Class=TableBorder>

  <tr>
    <th height="25" width="100%" colspan="5">产品列表</th>
  </tr>
  <tr height="25" style="font-size:12px;">
  <td width="40%" align="center">产品名称</td>
  <td width="15%" align="center">产品型号</td> 
  <td width="20%" align="center">上传时间</td>
  <td width="10%" align="center">点击次数</td>
  <td width="10%" align="center">选择</td></tr>
<%
dim page,sql,rs 
set rs=server.CreateObject("adodb.recordset")
if ClassId >0 then
 
sql="select * from products where ClassId=" & ClassId
rs.open sql,conn,1,1
else
 
sql="select * from products"
rs.open sql,conn,1,1
end if

if request("page")<>"" then  
    Page = cint(Request.QueryString("page"))                             'page值为接受值 
else
    Page=1
end if 
const PageSize = 20                                                      '每页显示记录数
rs.PageSize=PageSize

    if Page > rs.PageCount then                                          '如果接收的页数大于总页数
        Page = rs.PageCount                                              '设置当前显示页等于最后页        
    end if  
	if Page <= 0 then                                                    '如果page小于等于0
        Page = 1                                                         '设置PAGE等于第一页
    end if
        rs.AbsolutePage = Page                                           '如果大于零,显示当前页等于接收的页数 
		
if rs.eof then
response.write "<tr><td class=fl1>还没有任何产品！</td></tr>" 
else
   
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
    <td class=<%=flcss%> align="center">
	<a href="?action=edit&id=<%=rs("id")%>&classid=<%= rs("ClassId") %>"><%=rs("ProName")%></a></td>
    <td class=<%=flcss%> align="center"><%=rs("ProType")%></td>
    <td class=<%=flcss%> align="center"><%=rs("Time")%></td>
	<td class=<%=flcss%> align="center"><%=rs("Hits")%></td>
	
	
	<td class=<%=flcss%> align="center"><input name="id" type="checkbox" id="id" value="<%=rs("id")%>"></td>
  </tr>
  <%
          i=i+1
	      if i>=PageSize then exit do
	      rs.movenext
	      loop
  %>

 <tr> 
    <td colspan="5" Class=fl2> 
      <div align="right"> 
   
 <table width=""100%""><tr><td align=right>
 <tr><td>
<a href="admin_products.asp?page=1">首页</a> | 
<a href="admin_products.asp?page=<%= Page-1 %>">上一页</a> | 
<a href="admin_products.asp?page=<%= Page+1 %>">下一页</a> |
<a href="admin_products.asp?page=<%= rs.pagecount %>">尾页</a>
  
&nbsp;共<%= rs.RecordCount %>个

&nbsp;&nbsp;<input type="checkbox" name="chkall" onClick="CheckAll(this.form)">全选/取消&nbsp;
&nbsp;<INPUT TYPE="submit" value="删除" name="action" style="width:40pt;height:22" class="button">&nbsp;&nbsp;</td></tr></table>
</div></td>
  </tr>

  <%
    rs.close   
set rs=nothing
%>
<% end if %>
</table>
</form>
<%end sub %>




<%
sub del()                                                             '删除项

Dim idlist,idArr,uid
dim rs,sql,SucDelNum,i
dim arrUploadFiles
if Request("id")="" Then
		ErrMsg=ErrMsg+"<li>操作失败，请选择相关产品进行操作！"
		Call Error(ErrMsg)
		Exit Sub
End if

idlist=Request("id")
idArr=Split(idlist,",")
SucDelNum=0
for i = 0 to ubound(idarr)
	uid=Clng(idarr(i))
	set rs=server.CreateObject("adodb.recordset")
	sql="delete from products where id="& uid
	rs.open sql,conn,1,3
	
	SucDelNum=SucDelNum+1
next
call adminsuc("成功删除&nbsp;<font color=brown>"&SucDelNum&"</font>&nbsp;件产品！")
end sub %>



<% 
Sub Edit()                                                             '编辑项
dim id,sql,rs
id=trim(request("id"))
if id="" or not ISNumeric(id) then
Errmsg=Errmsg+"<br>"+"<li>操作失败，参数丢失！"
Call Error(ErrMsg)
exit sub
end if
id=clng(id)
set rs=server.CreateObject("adodb.recordset")
sql="select * from products where id="&id
rs.open sql,conn,1,1
if rs.eof then
set rs=nothing
Errmsg=Errmsg+"<br>"+"<li>操作失败，找不到此产品！"
Call Error(ErrMsg)
exit sub
end if
set rs=server.CreateObject("adodb.recordset")
sql="select * from products where id="&id
rs.open sql,conn,1,3

%>

<form action="admin_products.asp?action=EditSave" name="myform" method=post onSubmit="return checkadd();">
  <table cellpadding="2" cellspacing="1" border="0" width="80%" class=TableBorder align=center>
    <tr> 
      <th colspan=2 height=23>产品修改</th>
    </tr>
    <tr>
      <td colspan="2" Class=fl1 align="center"><table width="90%" height="21" border="0" cellpadding="3" cellspacing="0">
        
		<tr>
          <td>&nbsp;产品名称：<input name="ProName" type="text" id="ProName" value="<%= rs("ProName") %>" size="35" maxlength="40">
          &nbsp;* </td>
        </tr>
		<tr>
          <td>&nbsp;产品型号：<input name="ProType" type="text" id="ProType" value="" size="35" maxlength="30">
         </td>
        </tr>
		<tr><td><input name="pid" type="text" id="pid" value="<%= id %>" style="display:none;" ></td></tr>
		<% tt=request("classid")%>
		
		<tr>
          <td>&nbsp;产品类型：
              <select name="ClassId" size="1" id="ClassId2">
		   
                <option value="0">请选择</option>
			    <%
                    set rs2=server.CreateObject("adodb.recordset") 
			        sql="select * from classify"
			        rs2.open sql,conn,1,1
			        do while not rs2.eof
		        %>
                <option value="<%= rs2("classid") %>"><%= rs2("classname") %></option>
		        	 
		        <%  
                    rs2.movenext
			    	loop
			    	rs2.close
                %>  
            </select>
            &nbsp;*	
          </td>
		<script language="javascript">

            var obj = document.getElementById("ClassId2");
            obj.value = '<%= tt %>';
		</script>
        </tr>
        <tr>
          <td>&nbsp;产品添加：<input name="AddUser" type="text" id="AddUser" value="<%= rs("AddUser") %>" size="20" maxlength="30"></td>
        </tr>
        <tr>
          <td align="center"><b>产品详细信息</b></td>
        </tr>
        <tr>
          <td align="center"><textarea name="ProContent" id="editor_id" style="width:100%;height:550px;visibility:hidden;"><%=rs("ProContent")%></textarea>
		  </td>
        </tr>
       
        <tr>
          <td>产品图片：<input name="DefaultPicUrl" type="text" id="DefaultPicUrl" size="50" maxlength="200" value="<%= rs("ProPic") %>">
              用于产品预览显示 <br><iframe id=pic_file frameborder=0 src="uploadfiles.asp" width="100%" height="50" scrolling=no></iframe></td>
        </tr>
       
	
      </table>
      </td>
    </tr>
    
    <tr align="center"> 
      <td colspan="2" Class=fl2> <input type="submit" name="Submit" value="发 布" class=button> 
        &nbsp;&nbsp; <input type=button value="返 回" onClick="location.href = 'admin_products.asp'" class=button>      </td>
    </tr>
  </table>
</form>
<%
set rs=nothing
end sub
%>



</BODY></HTML>