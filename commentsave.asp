<!--#include file="conn.asp"-->

<%
    If request.form("CompanyName")="" Then
    Response.Write("<script language=""JavaScript"">alert(""������û���빫˾���ƣ��뷵�ؼ�飡��"");history.go(-1);</script>")
    response.end
    end if
    If request.form("Receiver")="" Then
    Response.Write("<script language=""JavaScript"">alert(""������û������ϵ�ˣ��뷵�ؼ�飡��"");history.go(-1);</script>")
    response.end
    end if
    If request.form("Phone")="" Then
    Response.Write("<script language=""JavaScript"">alert(""������û������ϵ�绰���뷵�ؼ�飡��"");history.go(-1);</script>")
    response.end
    end if
    If request.form("title")="" Then
    Response.Write("<script language=""JavaScript"">alert(""������û������⣬�뷵�ؼ�飡��"");history.go(-1);</script>")
    response.end
    end if
    If request.form("content")="" Then
    Response.Write("<script language=""JavaScript"">alert(""������û�����������ݣ��뷵�ؼ�飡��"");history.go(-1);</script>")
    response.end
    end if
   <!--  ɾ����β�ո�,�����ַ�������ȡ��ֵ -->
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
    <!--  ����һ�������ݶ�  -->
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
    <!--  д�����ݵ������ݶ�  -->
    rs.update
    rs.close
%>

<script language="JavaScript">alert("���Գɹ�����"); window.location.href = "comment.asp";</script>
