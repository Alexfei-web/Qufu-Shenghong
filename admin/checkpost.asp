
<% 
    function ChkPost()
	dim server_v1,server_v2
	chkpost=false
	server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))   '得到请求这个页面的页面的url
	server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))    '得到请求这个页面的计算机名字
	if mid(server_v1,8,len(server_v2))<>server_v2 then        '从一个字符串中取子字符串
		chkpost=false
	else
		chkpost=true
	End if
    End function 

    session.Timeout=20
%>
