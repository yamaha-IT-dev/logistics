<!--#INCLUDE FILE = "constant.asp" -->
<!--#INCLUDE FILE = "dbase/DB_adovbs.inc" -->
<!--#INCLUDE FILE = "dbase/DB_sqlerror.asp" -->
<!--#INCLUDE FILE = "dbase/DB_database.asp" -->
<!--#INCLUDE FILE = "tools/UTL_utilities.asp" -->
<%
Session("ConnectionTimeout") = 15
Session("CommandTimeout")    = 30

Dim ConnString, conn, DatabaseLocation

Sub OpenBaseDataBase()
	set conn=Server.CreateObject("ADODB.Connection")
	'conn.Provider = "ODBC DSN"
	conn.Open "DSN=as400;UID=edi;PWD=yma179"
End Sub

Sub CloseBaseDataBase()
	conn.close
	set conn = nothing
End Sub
%>