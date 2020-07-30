<!--#INCLUDE FILE="admin_checkuser.asp" -->
<!--#include file="checkpost.asp"-->

<!--#include file="css/admin_left.css"-->
<html>
<head>
    <title>管理页面</title>
    <meta http-equiv="refresh" content="180">
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312">

    <script language="javascript1.2">
        function showsubmenu(sid){
            whichEl = eval("submenu" + sid);
            if (whichEl.style.display == "none"){
                eval("submenu" + sid + ".style.display=\"\";");
            }
            else{
                eval("submenu" + sid + ".style.display=\"none\";");
            }
        }
    </script>
</head>
<% 
if ChkPost=false then
	'emsg="请不要从其它站点提交表"
    response.Redirect("login.asp?emsg=请不要从其它站点提交表")
    Response.End()
    End if
    dim AdminName
    AdminName=session("admin_username")
    if AdminName="" then
    response.redirect"login.asp" 
    end if
%>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <table cellspacing="0" cellpadding="0" width="100%" align="left" border="0">
        <tbody>
            <tr>
                <td valign="top">
                    <table cellspacing="0" cellpadding="0" width="158" align="center">
                        <tbody>
                            <tr>
                                <td valign="bottom" height="42">
                                    <img height="38"
                                    src="images/title.gif" width="158">
                                </td>
                            </tr>
                        </tbody>
                    </table>
                    <div style="height: 15px;"></div>

                    <!-----------------------------间隔 -------------------   -->
                    <table cellspacing="0" cellpadding="0" width="158" align="center">
                        <tbody>
                            <tr>
                                <td class="menu_title" onmouseover="this.className='menu_title2';"
                                    onmouseout="this.className='menu_title';"
                                    background="images/title_bg_quit.gif" height="25">
                                    <span>
                                        <a href="admin_contact.asp" target="main">管理首页</a>
                                        | <a href="logout.asp" target="_top"><b>退出</b></a>
                                    </span> 
                                </td>
                            </tr>
                            <tr>
                                <td class="sec_menu" height="50">
                                    <img height="20"
                                        src="images/bullet.gif" width="15" border="0">您的身份：系统管理员
			                        <br>
                                    <img height="20"
                                        src="images/bullet.gif" width="15" border="0"><a href="../index.asp" target="_blank">网站首页</a></td>
                            </tr>
                        </tbody>
                    </table>
                    <div style="height: 15px;"></div>
                    <!-----------------------------间隔 -------------------   -->
                    <table cellspacing="0" cellpadding="0" width="158" align="center">
                        <tbody>
                            <tr>
                                <td class="menu_title" id="menuTitle1"
                                    onmouseover="this.className='menu_title2';" onclick="showsubmenu(1)"
                                    onmouseout="this.className='menu_title';"
                                    background="images/admin_left_1.gif"
                                    height="25"><span>常规管理</span> </td>
                            </tr>
                            <tr>
                                <td id="submenu1">
                                    <div class="sec_menu" style="width: 158px">
                                        <table cellspacing="0" cellpadding="0" width="150" align="center">
                                            <tbody>
                                                <tr>
                                                    <td height="10"></td>
                                                </tr>
                                                <tr>
                                                    <td height="23">
                                                        <img height="20" alt=""
                                                            src="images/bullet.gif" width="15" border="0"><a
                                                                href="admin_admin.asp"
                                                                target="main">管理员管理</a></td>
                                                </tr>
                                                <tr>
                                                    <td height="23">
                                                        <img height="20" alt=""
                                                            src="images/bullet.gif" width="15" border="0"><a
                                                                href="admin_aboutus.asp"
                                                                target="main">关于我们管理</a></td>
                                                </tr>

                                                <tr>
                                                    <td height="23">
                                                        <img height="20" alt=""
                                                            src="images/bullet.gif" width="15" border="0"><a
                                                                href="admin_contact.asp"
                                                                target="main">联系我们管理</a></td>
                                                </tr>

                                                <tr>
                                                    <td height="23">
                                                        <img height="20" alt=""
                                                            src="images/bullet.gif" width="15" border="0"><a
                                                                href="admin_classify.asp"
                                                                target="main">产品分类管理</a></td>
                                                </tr>
                                                <tr>
                                                    <td height="10"></td>
                                                </tr>


                                                <tbody></tbody>
                                        </table>
                                    </div>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                    <div style="height: 15px;"></div>
                    <!-----------------------------间隔 -------------------   -->
                    <table cellspacing="0" cellpadding="0" width="158" align="center">
                        <tbody>
                            <tr>
                                <td class="menu_title" id="menuTitle1 "
                                    onmouseover="this.className='menu_title2';" onclick="showsubmenu(2)"
                                    onmouseout="this.className='menu_title';"
                                    background="images/admin_left_2.gif"
                                    height="25"><span>新闻管理</span> </td>
                            </tr>
                            <tr>
                                <td id="submenu2">
                                    <div class="sec_menu" style="width: 158px">
                                        <table cellspacing="0" cellpadding="0" width="150" align="center">
                                            <tbody>
                                                <tr>
                                                    <td height="10"></td>
                                                </tr>
                                                <tr>
                                                    <td height="23">
                                                        <img height="20" alt=""
                                                            src="images/bullet.gif" width="15" border="0"><a href="admin_news.asp?action=add" target="main">发布新闻</a></td>
                                                </tr>

                                                <tr>
                                                    <td height="23">
                                                        <img height="20" alt=""
                                                            src="images/bullet.gif" width="15" border="0"><a
                                                                href="admin_news.asp"
                                                                target="main">新闻管理</a></td>
                                                </tr>
                                                <tr>
                                                    <td height="10"></td>
                                                </tr>
                                                <tbody></tbody>
                                        </table>
                                    </div>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                    <div style="height: 15px;"></div>
                    <!-----------------------------间隔 -------------------   -->
                    <table cellspacing="0" cellpadding="0" width="158" align="center">
                        <tbody>
                            <tr>
                                <td class="menu_title" id="menuTitle1"
                                    onmouseover="this.className='menu_title2';" onclick="showsubmenu(3)"
                                    onmouseout="this.className='menu_title';"
                                    background="images/admin_left_3.gif"
                                    height="25"><span>产品管理</span> </td>
                            </tr>
                            <tr>
                                <td id="submenu3">
                                    <div class="sec_menu" style="width: 158px">
                                        <table cellspacing="0" cellpadding="0" width="150" align="center">
                                            <tbody>
                                                <tr>
                                                    <td height="10"></td>
                                                </tr>
                                                <tr>
                                                    <td height="23">
                                                        <img height="20" alt=""
                                                            src="images/bullet.gif" width="15" border="0"><a
                                                                href="admin_products.asp?action=add"
                                                                target="main">发布新产品</a></td>
                                                </tr>
                                                <tr>
                                                    <td height="23">
                                                        <img height="20" alt=""
                                                            src="images/bullet.gif" width="15" border="0"><a
                                                                href="admin_products.asp"
                                                                target="main">产品管理</a></td>
                                                </tr>
                                                <tr>
                                                    <td height="10"></td>
                                                </tr>

                                                <tbody></tbody>
                                        </table>
                                    </div>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                    <div style="height: 15px;"></div>
                    <!-----------------------------间隔 -------------------   -->

                    <table cellspacing="0" cellpadding="0" width="158" align="center">
                        <tbody>
                            <tr>
                                <td class="menu_title" id="menuTitle1"
                                    onmouseover="this.className='menu_title2';" onclick="showsubmenu(7)"
                                    onmouseout="this.className='menu_title';"
                                    background="images/admin_left_7.gif"
                                    height="25"><span>留言管理</span> </td>
                            </tr>
                            <tr>
                                <td id="submenu7">
                                    <div class="sec_menu" style="width: 158px">
                                        <table cellspacing="0" cellpadding="0" width="150" align="center">
                                            <tbody>
                                                <tr>
                                                    <td height="10"></td>
                                                </tr>

                                                <tr>
                                                    <td height="23">
                                                        <img height="20" alt=""
                                                            src="images/bullet.gif" width="15" border="0"><a
                                                                href="admin_comment.asp"
                                                                target="main">留言管理</a></td>
                                                </tr>
                                                <tr>
                                                    <td height="10"></td>
                                                </tr>
                                                <tbody></tbody>
                                        </table>
                                    </div>
                                    <div style="width: 158px">
                                        <table cellspacing="0" cellpadding="0" width="135" align="center">
                                            <tbody>
                                                <tr>
                                                    <td height="20"></td>
                                                </tr>
                                            </tbody>
                                        </table>
                                    </div>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                    &nbsp; 
            </tr>
        </tbody>
    </table>
</body>
</html>

