<!--#INCLUDE FILE="../inc/conn.asp" -->
<!--#INCLUDE FILE="../inc/pst.asp" -->

<%
Set fl = New fl_Class
dim AdminName
dim Errmsg,admin_flag
AdminName=fl.CheckStr(Trim(Session(fl.fl_sn&"_AdminName")))

Sub Error(ErrMsg)
%>
<br><br><br><br>
<table cellpadding=0 cellspacing=0 border=0	width=90% Class=TableBorder	align=center>
<tr>
	<td>
		<table cellpadding=5 cellspacing=1 border=0	width=100%>
		<tr	align="center">	
		  <th width="100%"><b>css</b></th>
		</tr>
		<tr>
			<td class=fl1><table	width="100%" height="100%" border="0" Class=LightTableBody>
				<tr><td width=80 align=center><img src="images/information.gif" align=absmiddle></td>
					<td valign=center><br><b>css：</b><br><br>
					<li>css<br><br>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr	align="center">	
		  <td width="100%" Class=fl2><input type=button onclick="javascript:history.go(-1)" value="确定(O)" style="width:65px" accesskey=o class=button>
		  </td>
		</tr>  
		</table>
	</td>
</tr>
</table><BR>
<%End Sub%>
<%
Sub adminsuc(byval sucmsg,byval comeurl)
dim endtime,dothetime
if comeurl="" or isnull(comeurl) then
comeurl=Request.ServerVariables("HTTP_REFERER")
end if
endtime=timer()
dothetime=cstr(int(( (endtime-Startime)*10000 )+0.5)/10)
	Response.Write "<br>"&_
		"<table cellpadding=5 cellspacing=1 border=0 align=center width=""100%""  Class=TableBorder>"&_
			"<tr align=center><th>css</th></tr>"&_
			"<tr><td class=""fl1"" height=80>&nbsp;&nbsp;<li>"&sucmsg&"<br>&nbsp;&nbsp;<li>共耗时 <font color=red>" & dothetime & " 毫秒</font><br>&nbsp;&nbsp;<li><a href="&comeurl&"><font color=red>css</font></a><br></td></tr>"&_
		"</table><BR>"
End Sub
%>