<link href="css/style.css" rel="stylesheet" type="text/css" />
    <link href="css/reset.css" rel="stylesheet" type="text/css" />
    <link href="css/home.css" rel="stylesheet" type="text/css" />
<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
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
	 
	 <div class="md_pinfo">

			 <div class="md_right_pinfo">	
						<div class="md_right_show">
						     <div class="box2">
							       <div class="title3_top">
								         <div class="ico3"><img src="images/hh_10.gif" /></div>
								         <div class="title3_wd"><a href="products.asp">产品信息</a></div>
							             
								   </div>
								   <div class="content3" style="text-align:center">
								          <div class="content311">
										       
											   <% 
											   id=request("id")
											   set rs=server.CreateObject("adodb.recordset")
									         sql="select * from products where id="&id
									         rs.open sql,conn,1,3
									        dim i
											 if rs("Hits")="" or rs("Hits")=null then
											  i=0
											  else
											   i=rs("Hits")
											  end if
											 rs("Hits")=i+1
											 rs.update
									         %>
									       <table cellpadding="0" cellspacing="0" border="0" style="margin: 0 auto;">
										        <tr><td style="text-align:center; padding-top:10px; padding-bottom:10px;"><img src="<%= rs("ProPic") %>" width="300" /></td></tr>
												
										         <tr><td style="font-size:12px; color:#333333; font-family:'宋体'; text-align:left; font-weight:bold;" height="25px;">产品名称：<%= rs("ProName") %></td></tr>
												 
												 <tr><td style="font-size:12px; color:#333333; font-family:'宋体'; text-align:left; font-weight:bold;" height="25px;">详细信息：<%= rs("ProType") %></td></tr>
												 <tr><td style="font-size:12px; color:#333333; text-align:left;  line-height:25px; text-indent:2em; "><%= rs("ProContent") %></td></tr>
										   </table>
									   
										  </div>
								          
								   </div>
							 </div>
						</div><!--  md_right_show end-->
						
						
			 </div>  <!--  md_right end-->
	 </div>  <!--  md end-->
	 <div style="clear:both; height:10px;"></div>
	 <!--#include file="footer.asp"--> <!--  footer end-->
	 
 </div>  <!--  main end-->
</body>
</html>
