<!--#include file="admin_Checkuser.asp"-->
<%Server.ScriptTimeOut=99999%>
<html>
<head>
    <title>����ҳ��</title>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312">
    <!--#include file="css/admin_main.css"-->
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <%
    admin_flag=",1,"
    If AdminName="" or AdminFlag="" Then
	    Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_Login.asp target=_top>��½</a>����롣"
	    Call Error(ErrMsg)
    else
    %>
    <table cellpadding="3" cellspacing="1" border="0" class="tableBorder" align="center">
        <tr>
            <th class="tableHeaderText" colspan="2" height="25"><%=fl.Setup(0)%>��Ϣͳ��</th>
            <tr>
                <td class="bodytitle" height="23" colspan="2">ϵͳ��Ϣ��δ��˵Ļ�Ա��<b><%=fl.execute("Select count(*) from Fl_Users where UserLock=1")(0)%></b>&nbsp;&nbsp;δ����Ķ�����<b><%=fl.execute("Select count(*) from Fl_Order where IsSend=0")(0)%></b></td>
            </tr>
        <tr>
            <td class="fl2" height="23" colspan="2">&nbsp;</td>
        </tr>
        <%if instr(","&AdminFlag&",",admin_flag)>0 then%>
        <tr>
            <td width="50%" class="fl1" height="23">���������ͣ� <%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</td>
            <td width="50%" class="fl1">�ű��������棺 <%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
        </tr>
        <tr>
            <td width="50%" class="fl1" height="23">վ������·���� <%=request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
            <td width="50%" class="fl1">IIS �汾�� <%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
        </tr>
        <tr>
            <td class="fl1" height="23">FSO�ļ���д�� 
                 <%
                     If fl.IsObjInstalled("Scripting.FileSystemObject") Then
	                    Response.Write "<font color=green><b>��</b></font>"
                     Else
	                    Response.Write "<font color=red><b>��</b></font>"
                     End If
                 %>    
            </td>
            <td class="fl1">Jmail�����ʼ�֧�֣� 
                <%
                    If fl.IsObjInstalled("JMail.SmtpMail") Then
                    	Response.Write "<font color=green><b>��</b></font>"
                    Else
                    	Response.Write "<font color=red><b>��</b></font>"
                    End If
                %>    

            </td>
        </tr>
        <tr>
            <td class="fl1" height="23">CDONTS�����ʼ�֧�֣� 
              <%
                If fl.IsObjInstalled("CDONTS.NewMail") Then
                	Response.Write "<font color=green><b>��</b></font>"
                Else
                	Response.Write "<font color=red><b>��</b></font>"
                End If
              %>    

            </td>
            <td class="fl1">AspEmail�����ʼ�֧�֣� 
                 <%
                   If fl.IsObjInstalled("Persits.MailSender") Then
                   	Response.Write "<font color=green><b>��</b></font>"
                   Else
                   	Response.Write "<font color=red><b>��</b></font>"
                   End If
                 %>    
            </td>
        </tr>
        <tr>
            <td class="fl1" height="23">������ϴ�֧�֣� 
                <%
                  If fl.IsObjInstalled("Adodb.Stream") Then
                  	Response.Write "<font color=green><b>��</b></font>"
                  Else
                  	Response.Write "<font color=red><b>��</b></font>"
                  End If
                %>    
            </td>
            <td class="fl1">AspUpload�ϴ�֧�֣� 
                <%
                  If fl.IsObjInstalled("Persits.Upload") Then
                  	Response.Write "<font color=green><b>��</b></font>"
                  Else
                  	Response.Write "<font color=red><b>��</b></font>"
                  End If
                %>    

            </td>
        </tr>
        <tr>
            <td width="50%" class="fl1" height="23">AspJpeg����Ԥ��ͼƬ��
                 <%
                   If fl.IsObjInstalled("Persits.Jpeg") Then
                   	Response.Write "<font color=green><b>��</b></font>"
                   Else
                   	Response.Write "<font color=red><b>��</b></font>"
                   End If
                 %>
            </td>
            <td width="50%" class="fl1">���ݿ�ʹ�ã�
                <%
                  If fl.IsObjInstalled("adodb.connection") Then
                  	Response.Write "<font color=green><b>��</b></font>"
                  Else
                  	Response.Write "<font color=red><b>��</b></font>"
                  End If
                %>
            </td>
        </tr>
        <form action="admin_Main.asp" method="post" id="form1" name="form1">
            <tr>
                <td height="23" colspan="2" class="fl1">�������֧�������ѯ��
                    <input class="input" type="text" value="" name="classname" size="30">
                    <input type="submit" value="�� ѯ" id="submit1" name="submit1">
                    ��������� ProgId �� ClassId&nbsp;&nbsp; 
                    <%
                        If Request("classname")<>"" Then
                        	If fl.IsObjInstalled(Request("classname")) Then
                        		Response.Write "<font color=green><b>��ϲ����������֧�� "&Request("classname")&" ���</b></font><BR>"
                        	Else
                        		Response.Write "<font color=red><b>��Ǹ������������֧�� "&Request("classname")&" ���</b></font><BR>"
                        	End If
                        End If
                    %>      
                </td>
            </tr>
        </form>
        <%Response.Flush%>
        <tr>
            <td height="11" colspan="2" class="fl2">
                <div align="center">
                    <b>ASP�ű����ͺ������ٶȲ���</b>
                </div>
            </td>
        </tr>
        <tr>
            <td height="11" colspan="2" class="fl1">
            <%

            	Response.Write "����������ԣ����ڽ���50��μӷ�����..."
            	dim lsabc,thetime,thetime2,t1,t2,i
            	t1=timer
            	for i=1 to 500000
            		lsabc= 1 + 1
            	next
            	t2=timer
            	thetime=cstr(int(( (t2-t1)*10000 )+0.5)/10)
            	Response.Write "...����ɣ�����ʱ <font color=red>" & thetime & " ����</font><br>"
            
            
            	Response.Write "����������ԣ����ڽ���20��ο�������..."
            	t1=timer
            	for i=1 to 200000
            		lsabc= 2^0.5
            	next
            	t2=timer
            	thetime2=cstr(int(( (t2-t1)*10000 )+0.5)/10)
            	Response.Write "...����ɣ�����ʱ <font color=red>" & thetime2 & " ����</font><br>"
            %>
            </td>
        </tr>
        <%end if%>
    </table>
    <br>
    <%Response.Flush%>
    <table cellpadding="3" cellspacing="1" border="0" class="tableBorder" align="center">
        <tr>
            <th colspan="2" height="25">ϵͳ�����ݷ�ʽ</th>
            <tr>
                <form method="POST" action="admin_user.asp">
                    <tr>
                        <td width="20%" class="fl1" height="23">���ٲ����û�</td>
                        <td width="80%" class="fl1">
                            <input type="text" name="keyword" size="30">
                            <input type="submit" value="���̲���">
                        </td>
                </form>
            </tr>
        <tr>
            <td width="20%" class="fl1" height="23">��ݹ�������</td>
            <td width="80%" class="fl1"><a href="admin_Order.asp">��Ʒ��������</a> | <a href="admin_book.asp">�ÿ����Թ���</a> | <a href="ReloadCache.asp">���·���������</a></td>
        </tr>
    </table>
    <br>
    <%
        End if
        set fl=nothing
        Call CloseDatabase()
    %>
</body>
</html>
