<!--#INCLUDE FILE="admin_checkuser.asp" -->
<!--#include file="checkpost.asp"-->

<!--#include file="css/admin_left.css"-->
<html>
<head>
    <title>����ҳ��</title>
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
	'emsg="�벻Ҫ������վ���ύ��"
    response.Redirect("login.asp?emsg=�벻Ҫ������վ���ύ��")
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

                    <!-----------------------------��� -------------------   -->
                    <table cellspacing="0" cellpadding="0" width="158" align="center">
                        <tbody>
                            <tr>
                                <td class="menu_title" onmouseover="this.className='menu_title2';"
                                    onmouseout="this.className='menu_title';"
                                    background="images/title_bg_quit.gif" height="25">
                                    <span>
                                        <a href="admin_contact.asp" target="main">������ҳ</a>
                                        | <a href="logout.asp" target="_top"><b>�˳�</b></a>
                                    </span> 
                                </td>
                            </tr>
                            <tr>
                                <td class="sec_menu" height="50">
                                    <img height="20"
                                        src="images/bullet.gif" width="15" border="0">������ݣ�ϵͳ����Ա
			                        <br>
                                    <img height="20"
                                        src="images/bullet.gif" width="15" border="0"><a href="../index.asp" target="_blank">��վ��ҳ</a></td>
                            </tr>
                        </tbody>
                    </table>
                    <div style="height: 15px;"></div>
                    <!-----------------------------��� -------------------   -->
                    <table cellspacing="0" cellpadding="0" width="158" align="center">
                        <tbody>
                            <tr>
                                <td class="menu_title" id="menuTitle1"
                                    onmouseover="this.className='menu_title2';" onclick="showsubmenu(1)"
                                    onmouseout="this.className='menu_title';"
                                    background="images/admin_left_1.gif"
                                    height="25"><span>�������</span> </td>
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
                                                                target="main">����Ա����</a></td>
                                                </tr>
                                                <tr>
                                                    <td height="23">
                                                        <img height="20" alt=""
                                                            src="images/bullet.gif" width="15" border="0"><a
                                                                href="admin_aboutus.asp"
                                                                target="main">�������ǹ���</a></td>
                                                </tr>

                                                <tr>
                                                    <td height="23">
                                                        <img height="20" alt=""
                                                            src="images/bullet.gif" width="15" border="0"><a
                                                                href="admin_contact.asp"
                                                                target="main">��ϵ���ǹ���</a></td>
                                                </tr>

                                                <tr>
                                                    <td height="23">
                                                        <img height="20" alt=""
                                                            src="images/bullet.gif" width="15" border="0"><a
                                                                href="admin_classify.asp"
                                                                target="main">��Ʒ�������</a></td>
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
                    <!-----------------------------��� -------------------   -->
                    <table cellspacing="0" cellpadding="0" width="158" align="center">
                        <tbody>
                            <tr>
                                <td class="menu_title" id="menuTitle1 "
                                    onmouseover="this.className='menu_title2';" onclick="showsubmenu(2)"
                                    onmouseout="this.className='menu_title';"
                                    background="images/admin_left_2.gif"
                                    height="25"><span>���Ź���</span> </td>
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
                                                            src="images/bullet.gif" width="15" border="0"><a href="admin_news.asp?action=add" target="main">��������</a></td>
                                                </tr>

                                                <tr>
                                                    <td height="23">
                                                        <img height="20" alt=""
                                                            src="images/bullet.gif" width="15" border="0"><a
                                                                href="admin_news.asp"
                                                                target="main">���Ź���</a></td>
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
                    <!-----------------------------��� -------------------   -->
                    <table cellspacing="0" cellpadding="0" width="158" align="center">
                        <tbody>
                            <tr>
                                <td class="menu_title" id="menuTitle1"
                                    onmouseover="this.className='menu_title2';" onclick="showsubmenu(3)"
                                    onmouseout="this.className='menu_title';"
                                    background="images/admin_left_3.gif"
                                    height="25"><span>��Ʒ����</span> </td>
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
                                                                target="main">�����²�Ʒ</a></td>
                                                </tr>
                                                <tr>
                                                    <td height="23">
                                                        <img height="20" alt=""
                                                            src="images/bullet.gif" width="15" border="0"><a
                                                                href="admin_products.asp"
                                                                target="main">��Ʒ����</a></td>
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
                    <!-----------------------------��� -------------------   -->

                    <table cellspacing="0" cellpadding="0" width="158" align="center">
                        <tbody>
                            <tr>
                                <td class="menu_title" id="menuTitle1"
                                    onmouseover="this.className='menu_title2';" onclick="showsubmenu(7)"
                                    onmouseout="this.className='menu_title';"
                                    background="images/admin_left_7.gif"
                                    height="25"><span>���Թ���</span> </td>
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
                                                                target="main">���Թ���</a></td>
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

