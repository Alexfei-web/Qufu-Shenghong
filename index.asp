<%@  language="VBSCRIPT" codepage="936" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
    <title>曲阜圣弘塑胶制品有限公司</title>
    <link rel="stylesheet" href="css/lightbox.css" type="text/css" media="screen" />
    <link href="css/style.css" rel="stylesheet" type="text/css" />
    <link href="css/reset.css" rel="stylesheet" type="text/css" />
    <link href="css/home.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="js/prototype.js"></script>
    <script type="text/javascript" src="js/scriptaculous.js?load=effects"></script>
    <script type="text/javascript" src="js/lightbox.js"></script>
    <script type="text/javascript" src="js/effects.js"></script>
</head>

<body>
    <!--#include file="conn.asp"-->
    <div id="main">
        <!--#include file="header.asp"-->
        <!--  header end-->
        <div style="clear: both;"></div>
        <div class="md">
            <div class="floor3">
                <div class="little-title">
                    <div class="left fl"></div>
                    <div class="right fl"></div>
                    <span><a href="products.asp" target="_blank">产品介绍</a></span>
                </div>
                <div class="container clearfix">
                    <div class="chanpin fl">
                        <a href="products.asp" target="_blank">
                            <img src="img/muju01.jpg">
                            <div class="text">
                                <div class="imgtext"><span>模具设计</span></div>
                            </div>
                        </a>
                    </div>
                    <div class="chanpin fl">
                        <a href="products.asp" target="_blank">
                            <img src="img/guan01.jpg">
                            <div class="text">
                                <div class="imgtext"><span>塑胶管</span></div>
                            </div>
                        </a>
                    </div>
                    <div class="chanpin fl">
                        <a href="products.asp" target="_blank">
                            <img src="img/dianzi01.jpg">
                            <div class="text">
                                <div class="imgtext"><span>电子产品</span></div>
                            </div>
                        </a>
                    </div>
                    <div class="chanpin fl">
                        <a href="products.asp" target="_blank">
                            <img src="img/daily01.jpg">
                            <div class="text">
                                <div class="imgtext"><span>塑料日用品</span></div>
                            </div>
                        </a>
                    </div>

                </div>
            </div>


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
                            <img src="img/jianjie.jpg">
                        </div>
                    </div>
                </div>


                <div class="md_right_right0">
                    <div class="little-title">
                        <div class="left fl"></div>
                        <div class="right fl"></div>
                        <span><a href="news.asp" target="_blank">新闻中心</a></span>
                    </div>
                    <div class="md_right_right">
                        <div class="box2">
                            <div class="title3_top">    
                                <div class="more">
                                    <a href="news.asp" target="_blank">更多</a>
                                </div>
                            </div>
                            <div class="content3">
                                <div class="content311">
                                    <table width="1200px" cellpadding="0" cellspacing="0">

                                        <!--  top 显示几个新闻-->
                                        <% set rs=server.CreateObject("adodb.recordset")
									  sql="select top 10 * from news order by id desc"
									  rs.open sql,conn,1,1
                                        %>
                                        <% 
									   if rs.eof and rs.bof then
									       response.write("<tr><td><p align='center'><font color='#ff0000'>还没任何动态</font></p></td></tr>")
									   else
									      do while not rs.eof 
                                        %>
                                        <tr>
                                            <td width="70%" height="50px" style="text-align: left; font-size: 18px;border-bottom:4px solid #fff ;padding: 0 50px;" class="rightnews_list"><a href="news_info.asp?id=<%= rs("id") %>" title="<%= rs("title") %>"><%= rs("title") %></a></td>
                                            <td width="30%" height="50px" style="text-align: right; color: #999999; font-size: 18px;border-bottom:4px solid #fff;padding: 0 10px;"><%= rs("time") %></td>
                                        </tr>
                                        <% 
										  rs.movenext
										  loop
										  end if
										  rs.close
                                        %>
                                    </table>
                                </div>
                            </div>
                        </div>


                    </div>
                </div>
                <!--  md_right_right end-->



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
