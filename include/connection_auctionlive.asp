<%
Session("ConnectionTimeout") = 15
Session("CommandTimeout")    = 30

Dim ConnString, conn, DatabaseLocation

Sub OpenDataBase()
	set conn=Server.CreateObject("ADODB.Connection")
	conn.Provider = "sqloledb"
	conn.Open "DSN=172.29.64.9;UID=webuser;PWD=w3bu53r;DATABASE=yamaha_auction"
End Sub

Sub CloseDataBase()
	conn.close
	set conn = nothing
End Sub

Sub OpenBaseDataBase()
	set conn=Server.CreateObject("ADODB.Connection")
	'conn.Provider = "ODBC DSN"
	conn.Open "DSN=as400;UID=edi;PWD=yma179"
End Sub

Sub CloseBaseDataBase()
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

Sub OpenEmailConnection()
	Set oMail = Server.CreateObject("CDO.Message")
	Set iConf = Server.CreateObject("CDO.Configuration")
	Set Flds = iConf.Fields
					
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.sendgrid.net"
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic clear text
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "yamahamusicau"
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "str0ppy@16"
	iConf.Fields.Update
End Sub

Sub CloseEmailConnection()
	Set oMail.Configuration = iConf
		oMail.To 		= emailTo
		oMail.Cc		= emailCc
		oMail.Bcc		= emailBcc
		oMail.From 		= emailFrom
		oMail.Subject 	= emailSubject
		oMail.TextBody 	= emailBodyText
		oMail.Send
			
	Set iConf = Nothing
	Set Flds = Nothing
End Sub
%>