<div class="fixed">
    <!-- ҳͷ -->
    <div class="head">
        <div class="container clearfix">
            <img src="img/GPS.jpg">
            <span><a href="">����</a></span>
            <span><a href="">[�л�����]</a></span>
            <% 
                set rs=server.CreateObject("adodb.recordset")
				sql="select * from contact "
				rs.open sql,conn,1,1				  
			%>
            <span><a href="">��ѷ���绰��</a><a href=""><%= rs("phone") %></a></span>
            <span><a href="">����</a></span>
        </div>
    </div>
    <!-- ���� -->
    <div class="floor1">
        <div class="container clearfix">
            <div class="left">
                <div class="fl">
                    <img src="img/logo.jpg"></div>
                <div class="text">
                    <span>����ʥ���ܽ���Ʒ���޹�˾</span>
                    <div class="text2 fl">Qufu Shenghong Plastic products Co., Ltd</div>
                </div>
            </div>
            <div class="right fr">
                <ul class="clearfix">
                    <li><a href="index.asp">��ҳ</a></li>
                    <li><a href="aboutus.asp">��������</a></li>
                    <li><a href="products.asp">��Ʒ����</a></li>
                    <li><a href="news.asp">��������</a></li>
                    <li><a href="comment.asp">��������</a></li>
                    <li><a href="contact.asp">��ϵ����</a></li>
                </ul>
            </div>
        </div>
    </div>
</div>
<!-- bannerͼ -->
<div class="floor2"></div>
