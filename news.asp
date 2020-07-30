<link href="css/style.css" rel="stylesheet" type="text/css" />
<link href="css/reset.css" rel="stylesheet" type="text/css" />
<link href="css/home.css" rel="stylesheet" type="text/css" />
<%@  language="VBSCRIPT" codepage="936" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
    <title>曲阜圣弘塑胶制品有限公司</title>
</head>

<body>
    <!--#include file="conn.asp"-->
    <div id="main">
        <!--#include file="header.asp"-->
        <!--  header end-->

        <div class="md">
            <div class="little-title">
                <div class="left fl"></div>
                <div class="right fl"></div>
                <span>新闻中心</span>
            </div>
            <div class="md_right_new" id="new002">
                <div class="md_right_show" style="width: 1000px;">
                    <div class="box2">
                        <div class="title3_top">
                           

                        </div>
                        <div class="content3">
                            <div class="news_list">
                                <table cellpadding="0" cellspacing="0" width="970px">
                                    <% set rs=server.CreateObject("adodb.recordset")
									  sql="select * from news order by id desc"
									  rs.open sql,conn,1,1 
                                    %>
                                    <% dim page
									   rs.pagesize=12  
									   if request("epage")<>"" then  
									    page = cint(Request.QueryString("epage"))
										else
										 page=1
										end if
									   if page<1 then
									     page=1
										end if
									   if page>rs.pagecount then
									     page=rs.pagecount
									   end if
									   rs.absolutepage=page
									    
										for i=0 to rs.pagesize-1 
                                            if rs.bof or rs.eof then exit for
											 
                                    %>
                                    <tr>
                                        <td height="30" width="25" align="center">
                                            <img src="images/arrow01.gif" />
                                        </td>
                                        <td class="news_list1" style="font-size: 15px;height:40px;">
                                            <a style="font-size: 15px;" href="news_info.asp?id=<%= rs("id") %>" title="<%= rs("title") %>"><%= rs("title") %></a>
                                        </td>
                                        <td style="font-size: 15px; color: #333333; text-align: right;height:40px;" width="80"><%= rs("time") %></td>
                                    </tr>
                                    <% 
										  rs.movenext
										  				 
										  next				  
                                    %>
                                </table>
                                <div class="page">
                                    <a href="news.asp?epage=1">首页</a>
                                    &nbsp;<a href="news.asp?epage=<%= page-1 %>">上一页</a>
                                    &nbsp;<a href="news.asp?epage=<%= page+1 %>">下一页</a>
                                    &nbsp;<a href="news.asp?epage=<%= rs.pagecount %>">末页</a>
                                    &nbsp;<a>共<%= page %>/<%= rs.pagecount %>页</a>
                                </div>
                                <% rs.close %>
                            </div>

                        </div>
                    </div>
                </div>
                <!--  md_right_show end-->


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
