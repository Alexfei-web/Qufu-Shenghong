<div class="md_left">
    <div class="box1">
        <div class="title1">
            <div class="ico">
                <img src="images/ico.png" />
            </div>
            <div class="product_title">产品中心</div>
        </div>

        <div class="product_menu">
            <% set rs=server.CreateObject("adodb.recordset") 
						  sql="select * from classify order by classid"
						  rs.open sql,conn,1,1
            %>
            <div class="product_list">
                <% 
							   if rs.eof and rs.bof then
							   response.Write("<a>还没有任何分类</a>")
							   else
							    do while not rs.eof
								
                %>
                <a href="products.asp?classid=<%= rs("classid") %>"><%= rs("classname") %></a>
                <img src="images/menu_line.gif" />

                <%  rs.movenext
										  loop
										  end if
										  rs.close 
                %>
            </div>
        </div>
    </div>

</div>
<!--  md_left end-->
