<%@ LANGUAGE = VBScript CodePage = 936%>
<%
Option Explicit
Response.Buffer = True
Dim Startime
Dim SqlNowString
Dim fl
Dim Conn,Db
Const fl_WebDir="../../../../../../../../"            
Const fl_Version="Version 3.0"                         
Const IsDeBug = 1           
Startime = Timer()

Const IsSqlDataBase = 0

If IsSqlDataBase = 1 Then

Const SqlDatabaseName = "FengLing_Web"
Const SqlPassword = "FengLing_Web"
Const SqlUsername = "FengLing_Web"
Const SqlLocalName = "(local)"
'================================================================================================================
SqlNowString = "GetDate()"
Else

Db = "../css/css.mdb"
'================================================================================================================
SqlNowString = "Now()"
End If

Sub ConnectionDatabase()
	Dim ConnStr
If IsSqlDataBase = 1 Then
		ConnStr = "Provider = Sqloledb; User ID = " & SqlUsername & "; Password = " & SqlPassword & "; Initial Catalog = " & SqlDatabaseName & "; Data Source = " & SqlLocalName & ";"
	Else
		ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(db)
End If
	On Error Resume Next
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.open ConnStr
	If Err Then
		err.Clear
		Set Conn = Nothing
		Response.Write "css"
		Response.End
	End If
End Sub

Sub CloseDatabase()
   If IsObject(Conn) Then
       Conn.Close
       Set Conn=Nothing
   End If
End Sub
%>