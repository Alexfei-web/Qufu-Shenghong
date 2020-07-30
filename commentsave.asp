<!--#include file="conn.asp"-->

<%
    If request.form("CompanyName")="" Then
    Response.Write("<script language=""JavaScript"">alert(""错误：您没输入公司名称，请返回检查！！"");history.go(-1);</script>")
    response.end
    end if
    If request.form("Receiver")="" Then
    Response.Write("<script language=""JavaScript"">alert(""错误：您没输入联系人，请返回检查！！"");history.go(-1);</script>")
    response.end
    end if
    If request.form("Phone")="" Then
    Response.Write("<script language=""JavaScript"">alert(""错误：您没输入联系电话，请返回检查！！"");history.go(-1);</script>")
    response.end
    end if
    If request.form("title")="" Then
    Response.Write("<script language=""JavaScript"">alert(""错误：您没输入标题，请返回检查！！"");history.go(-1);</script>")
    response.end
    end if
    If request.form("content")="" Then
    Response.Write("<script language=""JavaScript"">alert(""错误：您没输入留言内容，请返回检查！！"");history.go(-1);</script>")
    response.end
    end if
   <!--  删除首尾空格,返回字符串，获取表单值 -->
    Add=trim(request.form("Add"))
    
    Phone=trim(request.form("Phone"))
    Fax=trim(request.form("Fax"))
    email=trim(request.form("email"))
    
    
    if Add="" then
    	Add=null
    end if
    
    if Fax="" then
    	Fax=null
    end if
    if Email="" then
    	Email=null
    end if
    
    Set rs = Server.CreateObject("ADODB.Recordset")
    sql="select * from comment"
    rs.open sql,conn,1,3
    <!--  建立一条空数据段  -->
    rs.addnew
    
    rs("company")=trim(request.form("CompanyName"))
    rs("addr")=Add
    
    rs("receiver")=trim(request.form("Receiver"))
    rs("phone")=trim(request.form("Phone"))
    rs("mobile")=trim(request.form("Mobile"))
    rs("fax")=Fax
    rs("email")=email
    rs("title")=trim(request.form("title"))
    rs("content")=request.form("content")
    
    rs("time")=year(now) & "-" & month(now)  &"- " & day(now)
    <!--  写入数据到打开数据段  -->
    rs.update
    rs.close
%>

<script language="JavaScript">alert("留言成功！！"); window.location.href = "comment.asp";</script>
