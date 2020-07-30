<div class="fixed">
    <!-- 页头 -->
    <div class="head">
        <div class="container clearfix">
            <img src="img/GPS.jpg">
            <span><a href="">曲阜</a></span>
            <span><a href="">[切换城市]</a></span>
            <% 
                set rs=server.CreateObject("adodb.recordset")
				sql="select * from contact "
				rs.open sql,conn,1,1				  
			%>
            <span><a href="">免费服务电话：</a><a href=""><%= rs("phone") %></a></span>
            <span><a href="">分享到</a></span>
        </div>
    </div>
    <!-- 导航 -->
    <div class="floor1">
        <div class="container clearfix">
            <div class="left">
                <div class="fl">
                    <img src="img/logo.jpg"></div>
                <div class="text">
                    <span>曲阜圣弘塑胶制品有限公司</span>
                    <div class="text2 fl">Qufu Shenghong Plastic products Co., Ltd</div>
                </div>
            </div>
            <div class="right fr">
                <ul class="clearfix">
                    <li><a href="index.asp">首页</a></li>
                    <li><a href="aboutus.asp">关于我们</a></li>
                    <li><a href="products.asp">产品介绍</a></li>
                    <li><a href="news.asp">新闻中心</a></li>
                    <li><a href="comment.asp">在线留言</a></li>
                    <li><a href="contact.asp">联系我们</a></li>
                </ul>
            </div>
        </div>
    </div>
</div>
<!-- banner图 -->
<div class="floor2"></div>
