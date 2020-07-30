<!--#INCLUDE FILE="ser.asp" -->
<%
if AdminName="" then
response.redirect"admin_Login.asp"
else
call main()
end if
set fl=nothing
Call CloseDatabase()
sub main()
%>
<HTML><HEAD>
<title> </title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<!--#include file="images/admin_left.css"-->
<SCRIPT language=javascript1.2>
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
}
}
</SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 marginwidth="0" marginheight="0">
<TABLE cellSpacing=0 cellPadding=0 width="100%" align=left border=0>
  <TBODY>
  <TR>
    <TD vAlign=top><TABLE cellSpacing=0 cellPadding=0 width=158 align=center>
      <TBODY>
        <TR>
          <TD class=menu_title onMouseOver="this.className='menu_title2';" 
          onmouseout="this.className='menu_title';" 
          background=images/title_bg_quit.gif height=25><SPAN> <A 
            href="admin_login.asp?action=logout" 
            target=_top><B>exit</B></A></SPAN> </TD>
        </TR>
        <tr>
          <td class=sec_menu height="50">css管理菜单（已禁用）</td>
        </tr>
      </TBODY>
    </TABLE>
    </TR></TBODY></TABLE>
</BODY></HTML>
<%end sub%>
