<!-- ҳβ -->
		<div class="footer">
			<div class="container clearfix">
				<div class="right clearfix">
                    <% 
                    set rs=server.CreateObject("adodb.recordset")
				    sql="select * from contact "
					rs.open sql,conn,1,1				  
				    %>
					<div class="top fl">
						<span>��ϵ����</span>
						<br>
						<img src="img/phone.jpg">
						<span><%= rs("phone") %></span>
					</div>
					<div class="bot clearfix fl">
						<div class="fl">���䣺<%= rs("email") %></div>
						<div class="fl">���棺<%= rs("fax") %></div>
						<div class="fl">��ַ��<%= rs("addr") %></div>
					</div>
				</div>
			</div>
		</div>
		<!-- ��Ȩ -->
		<div class="banquan clearfix">
			<span>��Ȩ���У�ɽ���Ƽ���ѧGB1-307</span>
			<span>³ICP����000X00��</span>
		</div>