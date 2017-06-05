<!--#INCLUDE FILE = "constant.asp" -->
<!--#INCLUDE FILE = "functions-mpd.asp" -->
<!--#INCLUDE FILE = "dbase/DB_adovbs.inc" -->
<!--#INCLUDE FILE = "dbase/DB_sqlerror.asp" -->
<!--#INCLUDE FILE = "dbase/DB_database.asp" -->
<!--#INCLUDE FILE = "tools/UTL_utilities.asp" -->
<%
Session("ConnectionTimeout") = 15
Session("CommandTimeout")    = 30

'Session("strEmails")   = "Harsono_Setiono@gmx.yamaha.com" 
'Session("strSubject")  = "re-open record for "
'Session("strHosts")    = "smtp.yamahamusic.com"
'Session("strPort")     = "25"

Dim ConnString, conn, DatabaseLocation

Sub OpenDataBase()
	set conn=Server.CreateObject("ADODB.Connection")
	conn.Provider = "sqloledb"
	conn.Open "DSN=172.29.64.9;UID=webuser;PWD=w3bu53r;DATABASE=yamaha_mpd"
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