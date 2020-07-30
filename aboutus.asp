<%@  language="VBSCRIPT" codepage="936" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
    <title>曲阜圣弘塑胶制品有限公司</title>
    <link href="css/style.css" rel="stylesheet" type="text/css" />
    <link href="css/reset.css" rel="stylesheet" type="text/css" />
    <link href="css/home.css" rel="stylesheet" type="text/css" />
</head>

<body>
    <!--#include file="conn.asp"-->
    <div id="main">
        <!--#include file="header.asp"-->
        <!--  header end-->

        <div class="md">
            
            <div class="md_right">
                <div class="floor4">
                    <div class="little-title">
                        <div class="left fl"></div>
                        <div class="right fl"></div>
                        <span><a href="aboutus.asp" target="_blank">公司简介</a></span>
                    </div>
                    <div class="container clearfix">
                        <div class="left fl">
                             <div class="pg">
                                    <% set rs=server.CreateObject("adodb.recordset")
									  sql="select * from aboutus"
									  rs.open sql,conn,1,1
                                    %>
                                    <%= rs("content") %>
                                    <% rs.close %> 
                             </div>
                        </div>
                        <div class="pgg fr">
                            <img src="img/jianjie.jpg"></div>
                    </div>
                </div>

            </div>
            <!--  md_right end-->
        </div>
        <!--  md end-->
        <div style="clear: both; height: 10px;"></div>
        <!--#include file="footer.asp"-->
        <!--  footer end-->

    </div>
    <!--  main end-->
</body>
</html>
