<%@  language="vbscript" codepage="936" %>
<%
option explicit
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
'��Ҫ��ʹ������ֵ�ͼƬ�������
%>
<script language="javascript">
    function SetFocus() {
        if (document.Login.UserName.value == "")
            document.Login.UserName.focus();
        else
            document.Login.UserName.select();
    }
    function CheckForm() {
        if (document.Login.admin_username.value == "") {
            alert("�������û�����");
            document.Login.admin_username.focus();
            return false;
        }
        if (document.Login.admin_password.value == "") {
            alert("���������룡");
            document.Login.admin_password.focus();
            return false;
        }
        if (document.Login.admin_safecode.value == "") {
            alert("������������֤�룡");
            document.Login.admin_safecode.focus();
            return (false);
        }
    }
</script>
<html>
<head>
    <title>��½��վ</title>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312">
    <style type="text/css">
        INPUT {
            BORDER-RIGHT: #000000 1px solid;
            BORDER-TOP: #000000 1px solid;
            BORDER-LEFT: #000000 1px solid;
            BORDER-BOTTOM: #000000 1px solid;
            HEIGHT: 20px
        }

        TD {
            FONT-SIZE: 12px
        }

        .inll {
            BORDER-RIGHT: #ffffff;
            BORDER-TOP: #ffffff;
            BORDER-LEFT: #ffffff;
            BORDER-BOTTOM: #ffffff
        }

        .STYLE1 {
            color: #FFFFFF
        }

        .STYLE3 {
            color: #FFFFFF;
            font-weight: bold;
        }

        #login {
            background: url(Image/bac.jpg) no-repeat top center;
            background-size: cover;
        }
    </style>
</head>
<body id="login" bgcolor="#FFFFFF" background="" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td height="808" valign="middle">
                <table cellspacing="0" cellpadding="0" width="100%" align="center" border="0">
                    <tbody>
                        <tr>
                            <td height="160" align="center" valign="bottom">
                                <table width="59%" align="center" border="0">
                                    <tbody>
                                        <tr>
                                            <td width="17%" height="183" align="center" valign="middle"></td>

                                            <td width="83%">
                                                <form name="Login" action="logincheck.asp" method="post" target="_top" onsubmit="return CheckForm();">
                                                    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                                                        <tbody>
                                                            <tr height="40px">
                                                                <td width="31%" height="26" align="right" nowrap><span class="STYLE3">�û�����</span></td>
                                                                <td width="69%" height="26" nowrap>
                                                                    <input name="admin_username" type="text" id="admin_username" maxlength="20" style="width: 160px; border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1" onmouseover="this.style.background='#E1F4EE';" onmouseout="this.style.background='#FFFFFF'" onfocus="this.select(); "></td>
                                                            </tr>
                                                            <tr height="40px">
                                                                <td height="26" align="right" nowrap class="STYLE3">�ܡ��룺</td>
                                                                <td nowrap height="26">
                                                                    <input name="admin_password" type="password" maxlength="20" style="width: 160px; border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1" onmouseover="this.style.background='#E1F4EE';" onmouseout="this.style.background='#FFFFFF'" onfocus="this.select(); "></td>
                                                            </tr>
                                                            <tr height="40px">
                                                                <td height="26" align="right" nowrap class="STYLE3">��֤�룺</td>
                                                                <td nowrap height="26">
                                                                    <input name="admin_safecode" size="6" maxlength="4" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1" onmouseover="this.style.background='#E1F4EE';" onmouseout="this.style.background='#FFFFFF'" onfocus="this.select(); ">
                                                                    <img src="../inc/code.asp"></td>
                                                            </tr>
                                                            <tr height="40px">
                                                                <td></td>
                                                                <td height="26" align="left">
                                                                    <input type="submit" name="Submit" value=" ȷ&nbsp;�� " style="font-size: 9pt; height: 19; width: 60; color: #000000; background-color: #E1F4EE; border: 1 solid #E1F4EE" onmouseover="this.style.backgroundColor='#ffffff'" onmouseout="this.style.backgroundColor='#E1F4EE'">
                                                                    &nbsp;
                                                                    <input name="reset" type="reset" id="reset" value=" ��&nbsp;�� " style="font-size: 9pt; height: 19; width: 60; color: #000000; background-color: #E1F4EE; border: 1 solid #E1F4EE" onmouseover="this.style.backgroundColor='#ffffff'" onmouseout="this.style.backgroundColor='#E1F4EE'">
                                                                </td>
                                                            </tr>
                                                            
                                                        </tbody>
                                                        
                                                    </table>
                                                </form>
                                            </td>
                                        </tr>
                                    </tbody>
                                </table>
                                <div class="login001"><span style="margin: 10px 0; color: #333333;font-size:16px">�����û����Ȩ���벻Ҫ��¼�� </span></div>
                            </td>
                        </tr>
                    </tbody>

                </table>
            </td>
        </tr>

    </table>
    <script language="JavaScript" type="text/JavaScript">

        SetFocus();
    </script>
</body>
</html>
