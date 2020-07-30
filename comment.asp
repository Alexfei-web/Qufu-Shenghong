<link href="css/style.css" rel="stylesheet" type="text/css" />
<link href="css/reset.css" rel="stylesheet" type="text/css" />
<link href="css/home.css" rel="stylesheet" type="text/css" />
<%@  language="VBSCRIPT" codepage="936" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
    <title>曲阜圣弘塑胶制品有限公司</title>
    <script language="javaScript">
        function emailt() {
            if (document.form1.Email.value.length != 0) {
                if (document.form1.Email.value.charAt(0) == "." ||
                    document.form1.Email.value.charAt(0) == "@" ||
                    document.form1.Email.value.indexOf('@', 0) == -1 ||
                    document.form1.Email.value.indexOf('.', 0) == -1 ||
                    document.form1.Email.value.lastIndexOf("@") == document.form1.Email.value.length - 1 ||
                    document.form1.Email.value.lastIndexOf(".") == document.form1.Email.value.length - 1) {
                    alert("Email地址格式不正确！");
                    document.form1.Email.focus();
                    return false;
                }
            }
            else {
                alert("Email不能为空！");
                document.form1.Email.focus();
                return false;
            }
        }
    </script>
</head>

<body>
    <!--#include file="conn.asp"-->
    <div id="main">
        <!--#include file="header.asp"-->
        <!--  header end-->

        <div class="md">
            <div class="md_right">
                <div class="md_right_show">

                    <div class="box2">
                        <div class="little-title">
                            <div class="left fl"></div>
                            <div class="right fl"></div>
                            <span>在线留言</span>
                        </div>
                        <div class="content3">
                            <div style="height: 15px; clear: both;"></div>
                            <div class="">
                                <table cellpadding="0" cellspacing="3" border="0" align="center" width="94%">
                                    <tr>
                                        <td></td>
                                    </tr>
                                </table>
                            </div>
                            <table cellspacing="3" cellpadding="0" width="94%" align="center" border="0" style="font-size: 12px; font-family: '宋体';">
                                <tbody>
                                    <tr>
                                        <td>
                                            <form action="commentsave.asp" method="post" name="form1" id="form1" onsubmit="return emailt()">
                                                <table id="tab" width="700px" height="409" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#E3E3E3">
                                                    <tr>
                                                        <input type="hidden" name="Username" value="" />
                                                        <td height="30" align="right" bgcolor="#FFFFFF" style="font-size: 12px;"><span style="color: #FF0000">*</span>主题： </td>
                                                        <td height="30" bgcolor="#FFFFFF" style="padding-left: 10px;">
                                                        <input type="text" name="Title" size="42" maxlength="36" style="font-size: 12px" />
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td height="30" align="right" bgcolor="#FFFFFF" style="font-size: 12px;"><span style="color: #FF0000">*</span>内容：</td>
                                                        <td bgcolor="#FFFFFF" style="padding-left: 10px;">
                                                        <textarea rows="10" name="Content" cols="65" style="font-size: 12px"></textarea></td>
                                                    </tr>
                                                    <tr>
                                                        <td width="40%" height="30" align="right" bgcolor="#FFFFFF" style="font-size: 12px;"><span style="color: #FF0000">*</span>公司名称：</td>
                                                        <td width="60%" height="30" bgcolor="#FFFFFF" style="padding-left: 10px;">
                                                        <font><input type="text" name="CompanyName" size="30" maxlength="36" value="" style="font-size: 12px" /></font>
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td height="30" align="right" bgcolor="#FFFFFF" style="font-size: 12px;">公司地址：</td>
                                                        <td height="30" bgcolor="#FFFFFF" style="padding-left: 10px;"><font>
                                                        <input name="Add" type="text"  id="Add" style="font-size: 12px" value="" size="40" maxlength="60" />
                                                        </font></td>
                                                    </tr>

                                                    <tr>
                                                        <td height="30" align="right" bgcolor="#FFFFFF" style="font-size: 12px;"><span style="color: #FF0000">*</span>联系人：</td>
                                                        <td width="60%" height="-2" bgcolor="#FFFFFF" style="padding-left: 10px;"><font>
                                                        <input type="text" name="Receiver" size="12" maxlength="30" value="" style="font-size: 12px" />
                                                        </font> </td>
                                                    </tr>
                                                    <tr>
                                                        <td height="30" align="right" bgcolor="#FFFFFF" style="font-size: 12px;"><span style="color: #FF0000">*</span>联系电话：</td>
                                                        <td width="60%" height="-1" bgcolor="#FFFFFF" style="padding-left: 10px;"><font>
                                                        <input type="text" name="Phone" size="24" maxlength="36" value="" style="font-size: 12px" />
                                                        </font> </td>
                                                    </tr>
                                                    <tr>
                                                        <td height="30" align="right" bgcolor="#FFFFFF" style="font-size: 12px;">手机：</td>
                                                        <td height="11" bgcolor="#FFFFFF" style="padding-left: 10px;"><font>
                                                        <input name="Mobile" type="text" id="Mobile" style="font-size: 12px" value="" size="24" maxlength="36" />
                                                        </font></td>
                                                    </tr>
                                                    <tr>
                                                        <td height="30" align="right" bgcolor="#FFFFFF" style="font-size: 12px;">联系传真：</td>
                                                        <td width="60%" height="11" bgcolor="#FFFFFF" style="padding-left: 10px;"><font>
                                                        <input type="text" name="Fax" size="24" maxlength="36" value="" style="font-size: 12px" />
                                                        </font></td>
                                                    </tr>
                                                    <tr>
                                                        <td height="30" align="right" bgcolor="#FFFFFF" style="font-size: 12px;"><span style="color: #FF0000">*</span>E-mail：</td>
                                                        <td height="11" bgcolor="#FFFFFF" style="padding-left: 10px;"><font>
                                                        <input type="text" name="Email" size="28" maxlength="36" value="" style="font-size: 12px" />
                                                        </font></td>
                                                    </tr>
                                                    <tr>
                                                        <td width="50%" valign="top" bgcolor="#FFFFFF"></td>
                                                        <td width="50%" valign="middle" bgcolor="#FFFFFF" height="40" style="padding-left: 30px;">
                                                            <input type="submit" value="提交留言" name="cmdOk" />
                                                            <input type="reset" value="重写" name="cmdReset" />
                                                        </td>
                                                    </tr>
                                                </table>
                                            </form>
                                            <p>&nbsp;</p>
                                            <p></p>
                                        </td>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                        <!--  content3 end-->

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
