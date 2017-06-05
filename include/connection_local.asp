<%
Sub OpenDataBase()
	set conn=Server.CreateObject("ADODB.Connection")
	conn.Provider = "sqloledb"
	'conn.Open "DSN=intranet-sql;UID=webuser;PWD=w3bu53r;DATABASE=yamaha_it"
  conn.Open "Driver={SQL Server};Server=intranet-sql;Database=yamaha_it;User ID=webuser;Password=w3bu53r;"
End Sub

Sub CloseDataBase()
	conn.close
	set conn = nothing
End Sub
%>