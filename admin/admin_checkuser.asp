<!--#include file="conn.asp" -->
<%
Sub Error(ErrMsg)
%>
<br>
<br>
<br>
<br>
<table cellpadding="0" cellspacing="0" border="0" width="80%" class="TableBorder" align="center">
    <tr>
        <td>
            <table cellpadding="5" cellspacing="1" border="0" width="100%">
                <tr align="center">
                    <th width="100%"><b>��������-������Ϣ</b></th>
                </tr>
                <tr>
                    <td class="fl1">
                        <table width="100%" height="100%" border="0" class="LightTableBody">
                            <tr>
                                <td width="80" align="center">
                                    <img src="images/information.gif" align="absmiddle"></td>
                                <td valign="center">
                                    <br>
                                    <b>��������Ŀ���ԭ��</b><br>
                                    <br>
                                    <%= ErrMsg %><br>
                                    <br>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr align="center">
                    <td width="100%" class="fl2">
                        <input type="button" onclick="javascript: history.go(-1)" value="ȷ��(O)" style="width: 65px" accesskey="o" class="button">
                    </td>
                </tr>
            </table>
        </td>
    </tr>
</table>
<br>
<%End Sub%>
<%
    Sub adminsuc(byval sucmsg)    'byval��ֵ����
    dim endtime,dothetime
    if comeurl="" or isnull(comeurl) then
    comeurl=Request.ServerVariables("HTTP_REFERER")  '�����һ������ҳ��ĵ�ַ�������ж���ַ��Դ
    end if

	Response.Write "<br>"&_
	"<table cellpadding=5 cellspacing=1 border=0 align=center width=""100%""  Class=TableBorder>"&_
		"<tr align=center><th>�����ɹ�</th></tr>"&_
		"<tr><td class=""fl1"" height=80>&nbsp;&nbsp;<li>"&sucmsg&"<br>&nbsp;&nbsp;<li><a href="&comeurl&"><font color=red>������ﷵ��</font></a><br></td></tr>"&_
	"</table><BR>"
    End Sub
%>


<%
    Sub adminsuc2(byval sucmsg,byval comeurl)
    dim endtime,dothetime
    if comeurl="" or isnull(comeurl) then
    comeurl=Request.ServerVariables("HTTP_REFERER")
    end if

	Response.Write "<br>"&_
	"<table cellpadding=5 cellspacing=1 border=0 align=center width=""100%""  Class=TableBorder>"&_
		"<tr align=center><th>�����ɹ�</th></tr>"&_
		"<tr><td class=""fl1"" height=80>&nbsp;&nbsp;<li>"&sucmsg&"<br>&nbsp;&nbsp;<li><a href="&comeurl&"><font color=red>������ﷵ��</font></a><br></td></tr>"&_
	"</table><BR>"
    End Sub
%>