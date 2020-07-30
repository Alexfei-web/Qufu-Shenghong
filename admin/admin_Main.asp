<!--#include file="admin_Checkuser.asp"-->
<%Server.ScriptTimeOut=99999%>
<html>
<head>
    <title>管理页面</title>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312">
    <!--#include file="css/admin_main.css"-->
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <%
    admin_flag=",1,"
    If AdminName="" or AdminFlag="" Then
	    Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_Login.asp target=_top>登陆</a>后进入。"
	    Call Error(ErrMsg)
    else
    %>
    <table cellpadding="3" cellspacing="1" border="0" class="tableBorder" align="center">
        <tr>
            <th class="tableHeaderText" colspan="2" height="25"><%=fl.Setup(0)%>信息统计</th>
            <tr>
                <td class="bodytitle" height="23" colspan="2">系统信息：未审核的会员：<b><%=fl.execute("Select count(*) from Fl_Users where UserLock=1")(0)%></b>&nbsp;&nbsp;未处理的定单：<b><%=fl.execute("Select count(*) from Fl_Order where IsSend=0")(0)%></b></td>
            </tr>
        <tr>
            <td class="fl2" height="23" colspan="2">&nbsp;</td>
        </tr>
        <%if instr(","&AdminFlag&",",admin_flag)>0 then%>
        <tr>
            <td width="50%" class="fl1" height="23">服务器类型： <%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</td>
            <td width="50%" class="fl1">脚本解释引擎： <%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
        </tr>
        <tr>
            <td width="50%" class="fl1" height="23">站点物理路径： <%=request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
            <td width="50%" class="fl1">IIS 版本： <%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
        </tr>
        <tr>
            <td class="fl1" height="23">FSO文件读写： 
                 <%
                     If fl.IsObjInstalled("Scripting.FileSystemObject") Then
	                    Response.Write "<font color=green><b>√</b></font>"
                     Else
	                    Response.Write "<font color=red><b>×</b></font>"
                     End If
                 %>    
            </td>
            <td class="fl1">Jmail发送邮件支持： 
                <%
                    If fl.IsObjInstalled("JMail.SmtpMail") Then
                    	Response.Write "<font color=green><b>√</b></font>"
                    Else
                    	Response.Write "<font color=red><b>×</b></font>"
                    End If
                %>    

            </td>
        </tr>
        <tr>
            <td class="fl1" height="23">CDONTS发送邮件支持： 
              <%
                If fl.IsObjInstalled("CDONTS.NewMail") Then
                	Response.Write "<font color=green><b>√</b></font>"
                Else
                	Response.Write "<font color=red><b>×</b></font>"
                End If
              %>    

            </td>
            <td class="fl1">AspEmail发送邮件支持： 
                 <%
                   If fl.IsObjInstalled("Persits.MailSender") Then
                   	Response.Write "<font color=green><b>√</b></font>"
                   Else
                   	Response.Write "<font color=red><b>×</b></font>"
                   End If
                 %>    
            </td>
        </tr>
        <tr>
            <td class="fl1" height="23">无组件上传支持： 
                <%
                  If fl.IsObjInstalled("Adodb.Stream") Then
                  	Response.Write "<font color=green><b>√</b></font>"
                  Else
                  	Response.Write "<font color=red><b>×</b></font>"
                  End If
                %>    
            </td>
            <td class="fl1">AspUpload上传支持： 
                <%
                  If fl.IsObjInstalled("Persits.Upload") Then
                  	Response.Write "<font color=green><b>√</b></font>"
                  Else
                  	Response.Write "<font color=red><b>×</b></font>"
                  End If
                %>    

            </td>
        </tr>
        <tr>
            <td width="50%" class="fl1" height="23">AspJpeg生成预览图片：
                 <%
                   If fl.IsObjInstalled("Persits.Jpeg") Then
                   	Response.Write "<font color=green><b>√</b></font>"
                   Else
                   	Response.Write "<font color=red><b>×</b></font>"
                   End If
                 %>
            </td>
            <td width="50%" class="fl1">数据库使用：
                <%
                  If fl.IsObjInstalled("adodb.connection") Then
                  	Response.Write "<font color=green><b>√</b></font>"
                  Else
                  	Response.Write "<font color=red><b>×</b></font>"
                  End If
                %>
            </td>
        </tr>
        <form action="admin_Main.asp" method="post" id="form1" name="form1">
            <tr>
                <td height="23" colspan="2" class="fl1">其它组件支持情况查询：
                    <input class="input" type="text" value="" name="classname" size="30">
                    <input type="submit" value="查 询" id="submit1" name="submit1">
                    输入组件的 ProgId 或 ClassId&nbsp;&nbsp; 
                    <%
                        If Request("classname")<>"" Then
                        	If fl.IsObjInstalled(Request("classname")) Then
                        		Response.Write "<font color=green><b>恭喜，本服务器支持 "&Request("classname")&" 组件</b></font><BR>"
                        	Else
                        		Response.Write "<font color=red><b>抱歉，本服务器不支持 "&Request("classname")&" 组件</b></font><BR>"
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
                    <b>ASP脚本解释和运算速度测试</b>
                </div>
            </td>
        </tr>
        <tr>
            <td height="11" colspan="2" class="fl1">
            <%

            	Response.Write "整数运算测试，正在进行50万次加法运算..."
            	dim lsabc,thetime,thetime2,t1,t2,i
            	t1=timer
            	for i=1 to 500000
            		lsabc= 1 + 1
            	next
            	t2=timer
            	thetime=cstr(int(( (t2-t1)*10000 )+0.5)/10)
            	Response.Write "...已完成！共耗时 <font color=red>" & thetime & " 毫秒</font><br>"
            
            
            	Response.Write "浮点运算测试，正在进行20万次开方运算..."
            	t1=timer
            	for i=1 to 200000
            		lsabc= 2^0.5
            	next
            	t2=timer
            	thetime2=cstr(int(( (t2-t1)*10000 )+0.5)/10)
            	Response.Write "...已完成！共耗时 <font color=red>" & thetime2 & " 毫秒</font><br>"
            %>
            </td>
        </tr>
        <%end if%>
    </table>
    <br>
    <%Response.Flush%>
    <table cellpadding="3" cellspacing="1" border="0" class="tableBorder" align="center">
        <tr>
            <th colspan="2" height="25">系统管理快捷方式</th>
            <tr>
                <form method="POST" action="admin_user.asp">
                    <tr>
                        <td width="20%" class="fl1" height="23">快速查找用户</td>
                        <td width="80%" class="fl1">
                            <input type="text" name="keyword" size="30">
                            <input type="submit" value="立刻查找">
                        </td>
                </form>
            </tr>
        <tr>
            <td width="20%" class="fl1" height="23">快捷功能链接</td>
            <td width="80%" class="fl1"><a href="admin_Order.asp">产品定单管理</a> | <a href="admin_book.asp">访客留言管理</a> | <a href="ReloadCache.asp">更新服务器缓存</a></td>
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
