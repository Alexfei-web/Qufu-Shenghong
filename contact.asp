<link href="css/style.css" rel="stylesheet" type="text/css" />
<link href="css/reset.css" rel="stylesheet" type="text/css" />
<link href="css/home.css" rel="stylesheet" type="text/css" />
<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
    <title>����ʥ���ܽ���Ʒ���޹�˾</title>
</head>

<body>
<!--#include file="conn.asp"-->
<div id="main">
     <!--#include file="header.asp"-->
    <!--  header end-->
	 <div class="connection">
			<div class="container">
				<div class="text">
					<div class="text1">��ϵ��ʽ</div>
                    <% 
                        set rs=server.CreateObject("adodb.recordset")
						sql="select * from contact "
						rs.open sql,conn,1,1				  
					%>
					<div class="text2 clearfix">
						<div class="left fl">
							<div class="fl"><img src="img/gps.png" ></div>
                            <div><span>��˾��ַ��</span></div>
                            <div><span><%= rs("addr") %></span></div>
							
						</div>
						<div class="right fl">
							<div class="fl"><img src="img/message2.png" ></div>
                            <div><span>��ϵ�绰��</span></div>
                            <div><span><%= rs("phone") %></span></div>
						</div>
					</div>
					<div class="text2 clearfix">
						<div class="left fl">
							<div class="fl"><img src="img/message.png" ></div>
                            <div><span>���棺</span></div>
                            <div><span><%= rs("fax") %></span></div>
						</div>
						<div class="right fl">
							<div class="fl"><img src="img/xin.png" ></div>
                            <div><span>���䣺</span></div>
                            <div><span><%= rs("email") %></span></div>
						</div>
					</div>
                    <div class="text2 clearfix">
						<div class="left fl">
							<div class="fl"><img src="img/message.png" ></div>
                            <div><span>�ʱࣺ</span></div>
                            <div><span><%= rs("yb") %></span></div>
						</div>
						<div class="right fl">
							<div class="fl"><img src="img/message2.png" ></div>
                            <div><span>����绰��</span></div>
                            <div><span><%= rs("sj") %></span></div>
						</div>
					</div>
				</div>
				<div class="map"></div>
			</div>
		</div>

	 <!--#include file="footer.asp"--> <!--  footer end-->
 </div>  <!--  main end-->
</body>
</html>
