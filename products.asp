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

        <div class="md_products">
                <div class="little-title">
                    <div class="left fl"></div>
                    <div class="right fl"></div>
                    <span>产品介绍</span>
                </div>

            <div class="content3">
                            <div id="js" style="margin-left: 16px; overflow: hidden; width: 950px">
                                <table style="padding-left: 2px;" cellspacing="0" cellpadding="0" width="100%" border="0">
                                    <tr>
                                        <td id="js1" valign="top">
                                            <table height="100" width="1000" cellspacing="0" cellpadding="1" align="left" border="0">
                                                <tr>
                                                    <td><a href="pic_show/01.jpg" rel="lightbox">
                                                        <img src="pic_show/01.jpg" /></a></td>
                                                    <td width="2"></td>
                                                    <td><a href="pic_show/02.jpg" rel="lightbox">
                                                        <img src="pic_show/02.jpg" /></a></td>
                                                    <td width="2"></td>
                                                    <td><a href="pic_show/03.jpg" rel="lightbox">
                                                        <img src="pic_show/03.jpg" /></a></td>
                                                    <td width="2"></td>
                                                    <td><a href="pic_show/04.jpg" rel="lightbox">
                                                        <img src="pic_show/04.jpg" /></a></td>
                                                    <td width="2"></td>
                                                    <td><a href="pic_show/05.jpg" rel="lightbox">
                                                        <img src="pic_show/05.jpg" /></a></td>
                                                    <td width="2"></td>
                                                    <td><a href="pic_show/06.jpg" rel="lightbox">
                                                        <img src="pic_show/06.jpg" /></a></td>
                                                    <td width="2"></td>
                                                    <td><a href="pic_show/07.jpg" rel="lightbox">
                                                        <img src="pic_show/07.jpg" /></a></td>
                                                    <td width="2"></td>
                                                    <td><a href="pic_show/08.jpg" rel="lightbox">
                                                        <img src="pic_show/08.jpg" /></a></td>
                                                    <td width="2"></td>
                                                    <td><a href="pic_show/09.jpg" rel="lightbox">
                                                        <img src="pic_show/09.jpg" /></a></td>

                                                </tr>
                                            </table>
                                        </td>
                                        <td id="js2" valign="top"></td>
                                    </tr>
                                </table>
                                <script>
                                    var speed=30;
                                    js2.innerHTML=js1.innerHTML;
                                    function Marquee(){
                                        if(js2.offsetWidth-js.scrollLeft<=0)
                                        js.scrollLeft-=js1.offsetWidth
                                        else{
                                            js.scrollLeft++
                                       }
                                    }
                                    var MyMar=setInterval(Marquee,speed)
                                    js.onmouseover=function() {clearInterval(MyMar)}
                                    js.onmouseout=function() {MyMar=setInterval(Marquee,speed)}
                                </script>
                            </div>
                        </div>


            <!--#include file="left.asp"-->
            <!--  md_left end-->
            <div class="md_right">
                <div class="md_right_show">
                    <div class="box2">
                        <div class="title3_top">
                            <div class="ico3">
                                <img src="images/hh_10.gif" />
                            </div>
                            <div class="title3_wd"><a href="products.asp">产品中心</a></div>

                        </div>
                      
                        <div class="content3">
                            <div class="pic_list">
                                <% dim classid
											   classid=trim(request("classid"))
											   set rs=server.CreateObject("adodb.recordset")
											    if classid="" then
											   
									  sql="select * from products order by id desc"
									  rs.open sql,conn,1,1
									  else
									   sql="select * from products where ClassId=" & classid & " order by id desc"
									  rs.open sql,conn,1,1
									  end if
                                %>

                                <!--<script language="javascript">alert("")</script> -->

                                <% if rs.eof then
                                %>
                                <div class="productone">没有该类产品！</div>
                                <% else %>

                                <% dim page
										dim j
										j=0
									   rs.pagesize=12  
									   if request("epage")<>"" then  
									    page = cint(Request.QueryString("tpage"))
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



                                <div class="product_one">

                                    <a href="product_info.asp?id=<%= rs("id") %>" target="_blank">
                                        <img src="<%= rs("ProPic") %>"></a>
                                    <ul class="product_info">
                                        <li>名称：<a href="product_info.asp?id=<%= rs("id")%>"><%= rs("ProName") %></a></li>


                                    </ul>

                                </div>
                                <!--  单个产品信息product_one end-->

                                <% j=j+1 
											   if j mod 4=0 then%>
                                <div style="clear: both;"></div>
                                <%
											  
											    end if  %>

                                <% 
										  rs.movenext
										  
										 
										  next
										  
										  
                                %>



                                <div style="clear: both;"></div>
                                <div class="page" style="margin-bottom: 10px;">
                                    <a href="products.asp?tpage=1>&classid=<%=classid%>">首页</a>
                                    &nbsp;<a href="products.asp?tpage=<%= page-1 %>&classid=<%=classid%>">上一页</a>
                                    &nbsp;<a href="products.asp?tpage=<%= page+1 %>&classid=<%=classid%>">下一页</a>
                                    &nbsp;<a href="products.asp?tpage=<%= rs.pagecount %>&classid=<%=classid%>">末页</a>
                                    &nbsp;<a>共<%= page %>/<%= rs.pagecount %>页</a>
                                </div>


                                <% end if %>
                            </div>
                        </div>
                        <!--  产品列表content3 end-->
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
