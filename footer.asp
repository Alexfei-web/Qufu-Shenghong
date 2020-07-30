<!-- 页尾 -->
		<div class="footer">
			<div class="container clearfix">
				<div class="right clearfix">
                    <% 
                    set rs=server.CreateObject("adodb.recordset")
				    sql="select * from contact "
					rs.open sql,conn,1,1				  
				    %>
					<div class="top fl">
						<span>联系我们</span>
						<br>
						<img src="img/phone.jpg">
						<span><%= rs("phone") %></span>
					</div>
					<div class="bot clearfix fl">
						<div class="fl">邮箱：<%= rs("email") %></div>
						<div class="fl">传真：<%= rs("fax") %></div>
						<div class="fl">地址：<%= rs("addr") %></div>
					</div>
				</div>
			</div>
		</div>
		<!-- 版权 -->
		<div class="banquan clearfix">
			<span>版权所有：山东科技大学GB1-307</span>
			<span>鲁ICP备：000X00号</span>
		</div>