<%
	dim conn                   '����һ���������������ݿ�bai����
	dim connstr                
	dim db
	db="data/data_info.mdb"
	Set conn = Server.CreateObject("ADODB.Connection")     '�������ӱ���ΪADODB.Connection��������VB���ڽ������ݿ����ӵĶ���
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)  '��ָ�������ݿ�
	conn.Open connstr
	
	sub closedatabase()
	  conn.close
        set conn=Nothing
    End sub
%>
