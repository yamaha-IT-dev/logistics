<!--#INCLUDE FILE = "constant.asp" -->
<!--#INCLUDE FILE = "dbase/DB_adovbs.inc" -->
<!--#INCLUDE FILE = "dbase/DB_sqlerror.asp" -->
<!--#INCLUDE FILE = "dbase/DB_database.asp" -->
<!--#INCLUDE FILE = "tools/UTL_utilities.asp" -->
<%
Session("ConnectionTimeout") = 15
Session("CommandTimeout")    = 30

Dim ConnString, conn, DatabaseLocation
Dim local_address 

local_address = Request.ServerVariables("LOCAL_ADDR")

Sub OpenDataBase()
	'if local_address = "intranet" then
	'	set conn=Server.CreateObject("ADODB.Connection")
	'	conn.Provider = "sqloledb"
	'	conn.Open "DSN=172.29.64.9;UID=webuser;PWD=w3bu53r;DATABASE=YMADEV"
	'else
		set conn=Server.CreateObject("ADODB.Connection")
		conn.Open = "Provider=sqloledb; Network Library=DBMSSOCN; Data Source=wic007q.server-sql.com,4656; Initial Catalog=vs130299_1; User ID=vs130299_1; Password=ZfD4LTaiD6;"
	'end if
End Sub

Sub CloseDataBase()
	conn.close
	set conn = nothing
End Sub

'********************************************************************
'Function: FRM_BuildOptionList
'Description: This function builds an option list from a list of supplied
'options and a selected option.
'********************************************************************	
Sub FRM_BuildOptionList(strOptionList,strSelectedOption)
	Dim arrOptionList, strCurrentOption, strSelected
	Dim intLoop

	arrOptionList = Split(strOptionList,",",-1)

	for intLoop = 0 to Ubound(arrOptionList)
		strCurrentOption = arrOptionList(intLoop)
		strSelected = ""
		if strCurrentOption = strSelectedOption then
			strSelected = "selected"
		end if
		response.write("<option " & strSelected & " value='" & strCurrentOption & "'>" & strCurrentOption & "</option>")
	next
End Sub
%>