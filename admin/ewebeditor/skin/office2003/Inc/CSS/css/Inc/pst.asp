<%
Class fl_Class
	Public SqlQueryNum,fl_sn,CacheName,Reloadtime
	Public Setup
	Public UserName,UserID,UserInfo,UserTrueIP,Actforip
	Private Sub Class_Initialize() 
		fl_sn="FengLing_System"
		CacheName="FengLing_System"
		If Not Response.IsClientConnected Then
			Session(fl_sn &"_UserName")=empty
			Session(fl_sn &"_UserID")=empty
			Set fl=Nothing
			Response.End
		End If
		SqlQueryNum = 0           
		Reloadtime=30		
		UserName = CheckStr(Trim(Session(fl_sn &"_UserName")))
		UserID = Trim(Session(fl_sn &"_UserID"))
		If IsNumeric(UserID) = 0 Or UserID="" Then UserID=0
		UserID = Clng(UserID)
		UserInfo=split(Session(fl_sn &"_UserInfo"),"|")
		UserTrueIP = GetIP()
		Call IsloadCache()   
		Setup=Application(CacheName & "_Setup")
	End Sub
    
	'取得用户IP
	Private Function GetIP()
		Dim strIPAddr 
		If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" OR InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then 
			strIPAddr = Request.ServerVariables("REMOTE_ADDR") 
		ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then 
			strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",")-1) 
			Actforip=Request.ServerVariables("REMOTE_ADDR")
		ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then 
			strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";")-1)
			Actforip=Request.ServerVariables("REMOTE_ADDR")
		Else 
			strIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
			Actforip=Request.ServerVariables("REMOTE_ADDR")
		End If 
		GetIP = CheckStr(Trim(Mid(strIPAddr, 1, 30)))
	End Function

'分类列表
Public Sub ClassList(ChannelID,CID,SID)
dim sqlClass,rsClass,i,iDepth,TempStr
sqlClass="select * From Fl_Class where ParentID="&Clng(CID)&" And ChannelID="&Clng(ChannelID)&" order by OrderID"
set rsClass=Execute(sqlClass)
If not rsClass.eof then
do until rsClass.eof
iDepth=rsClass("Depth")
TempStr="<option value=" & rsClass("ClassID") & ""
if rsClass("ClassID")=SID Then TempStr=TempStr+" Selected"
TempStr=TempStr+ ">"
if iDepth>0 then 
for i=1 to iDepth
TempStr=TempStr+ "│"
TempStr=TempStr+"&nbsp;&nbsp;"
next 
end if

if rsClass("Child")>0 then 
TempStr=TempStr+"├" 
else 
TempStr=TempStr+ "└" 
end if
TempStr=TempStr+ "" & rsClass("ClassName") & "" 
if rsClass("Child")>0 then 
TempStr=TempStr+ "（" & rsClass("Child") & "）" 
end if
TempStr=TempStr+"</option>"
Response.Write(TempStr)
Call ClassList(ChannelID,rsClass("ClassID"),SID)
rsClass.movenext 
loop
else
exit sub
End if
set rsClass=nothing 
End Sub

	Public Function Checkstr(Str)
		If Isnull(Str) Then
			CheckStr = ""
			Exit Function 
		End If
		Str = Replace(Str,Chr(0),"")
		CheckStr = Replace(Str,"'","''")
	End Function
	Private Function IsloadCache()
		If Not(IsArray(Application(CacheName & "_Setup"))) or not(isdate(Application(CacheName & "_Setup_Updatetime"))) Then
		    Call loadSetup()
		ElseIf Reloadtime>0 Then
		    If DateDiff("s",CDate(Application(CacheName & "_Setup_Updatetime")),Now()) >= (60*Reloadtime) Then
		       Call loadSetup()
			End If
		End If
	End Function

	Private Sub loadSetup()
		Dim Rs
		Dim SetupStr
		Set Rs = Execute("Select Setting From [Fl_Setup]")
		if Rs.Bof or Rs.eof Then
		set fl=nothing
		set Rs=nothing
		Call CloseDatabase()
		response.Write("www")
		response.End
		end if
		SetupStr=split(Rs(0),"|||")
		set rs=nothing
        Application.Lock
        Application(CacheName & "_Setup")=SetupStr
        Application(CacheName & "_Setup_Updatetime")=Now()
        Application.UnLock
	End Sub

	Public Function ChkPost()
		Dim server_v1,server_v2
		Chkpost=False 
		server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
		server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
		If Mid(server_v1,8,len(server_v2))=server_v2 Then Chkpost=True 
	End Function
	Public Function Replacehtml(Textstr)
		Dim Str,re
		Str=Textstr
		Set re=new RegExp
			re.IgnoreCase =True
			re.Global=True
			re.Pattern="<(.[^>]*)>"
			Str=re.Replace(Str, "")
		Set Re=Nothing
			Replacehtml=Str
	End Function

	Public Function Execute(Command)
		If Not IsObject(Conn) Then ConnectionDatabase
		If IsDeBug = 0 Then 
			On Error Resume Next
			Set Execute = Conn.Execute(Command)
			If Err Then
				err.Clear
				Set Conn = Nothing
				Response.Write "www。"
				Response.End
			End If
		Else
			Set Execute = Conn.Execute(Command)
		End If	
		SqlQueryNum = SqlQueryNum+1
	End Function
	Public Function GetCode()
			GetCode= "<img src="""&fl_WebDir&"inc/GetCode.asp"" alt= ""wwcode"" style=""cursor : pointer;"" onclick=""this.src='"&fl_WebDir&"inc/GetCode.asp'""/> "
	End Function

	Public Function CodeIsTrue()
		Dim CodeStr
		CodeStr=Lcase(Trim(Request("CodeStr")))
		If CStr(Session("GetCode"))=CStr(CodeStr) And CodeStr<>""  Then
			CodeIsTrue=True
			Session("GetCode")=empty
		Else
			CodeIsTrue=False
			Session("GetCode")=empty
		End If	
	End Function
	Public Function HTMLEncode(fString)
		If Not IsNull(fString) Then
			fString = replace(fString, ">", "&gt;")
			fString = replace(fString, "<", "&lt;")
			fString = Replace(fString, CHR(32), "&nbsp;")		'&nbsp;
			fString = Replace(fString, CHR(9), " ")			'&nbsp;
			fString = Replace(fString, CHR(34), "&quot;")
			'fString = Replace(fString, CHR(39), "&#39;")	'单引号过滤
			fString = Replace(fString, CHR(13), "")
			fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
			fString = Replace(fString, CHR(10), "<BR> ")
			HTMLEncode = fString
		End If
	End Function

	Public Function strLength(str)
		If isNull(str) Or Str = "" Then
			StrLength = 0
			Exit Function
		End If
		Dim WINNT_CHINESE
		WINNT_CHINESE=(len("例子")=2)
		If WINNT_CHINESE Then
			Dim l,t,c
			Dim i
			l=len(str)
			t=l
			For i=1 To l
				c=asc(mid(str,i,1))
				If c<0 Then c=c+65536
				If c>255 Then t=t+1
			Next
			strLength=t
		Else 
			strLength=len(str)
		End If
	End Function
Public function gotTopic(str,strlen)
	if str="" then
		gotTopic=""
		exit function
	end if
	dim l,t,c, i
	l=len(str)
	t=0
	for i=1 to l
		c=Abs(Asc(Mid(str,i,1)))
		if c>255 then
			t=t+2
		else
			t=t+1
		end if
		if t>=strlen then
			gotTopic=left(str,i) & "…"
			exit for
		else
			gotTopic=str
		end if
	next
	gotTopic=gotTopic
end function

	
End Class
%>