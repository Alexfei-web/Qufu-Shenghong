<% 
    <!--  ����һ�����ݿ�����ʵ��  -->
    set conn=server.CreateObject("adodb.connection")
    conn.connectionstring="Provider= microsoft.jet.oledb.4.0; data source=" & server.MapPath("admin/data/data_info.mdb")
    conn.open
 %>