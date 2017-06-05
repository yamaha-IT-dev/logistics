<%
Dim conn

Sub OpenDataBase()
	set conn=Server.CreateObject("ADODB.Connection")
	conn.Provider = "sqloledb"
	conn.Open "DSN=172.29.64.9;UID=webuser;PWD=w3bu53r;DATABASE=YMADEV"
	
	set conn=Server.CreateObject("ADODB.Connection")
	conn.Provider = "sqloledb"
	conn.Open = "DSN=172.29.64.9;UID=webuser;PWD=w3bu53r;DATABASE=yamaha_it"
End Sub

Sub CloseDataBase()
	conn.close
	set conn = nothing
End Sub
%>