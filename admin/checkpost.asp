
<% 
    function ChkPost()
	dim server_v1,server_v2
	chkpost=false
	server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))   '�õ��������ҳ���ҳ���url
	server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))    '�õ��������ҳ��ļ��������
	if mid(server_v1,8,len(server_v2))<>server_v2 then        '��һ���ַ�����ȡ���ַ���
		chkpost=false
	else
		chkpost=true
	End if
    End function 

    session.Timeout=20
%>
