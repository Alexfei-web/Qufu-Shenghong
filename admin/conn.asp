<%
	dim conn                   '定义一个变量，用于数据库bai连接
	dim connstr                
	dim db
	db="data/data_info.mdb"
	Set conn = Server.CreateObject("ADODB.Connection")     '设置连接变量为ADODB.Connection对象，这是VB用于进行数据库连接的对象
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)  '打开指定的数据库
	conn.Open connstr
	
	sub closedatabase()
	  conn.close
        set conn=Nothing
    End sub
%>
