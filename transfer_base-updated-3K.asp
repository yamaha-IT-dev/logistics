<!--#include file="include/connection_it.asp " -->
<% strSection = "transfer" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Freight - Send Transfer BASE updated Email Notification</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/validations.js"></script>
<%
'-----------------------------------------------------------------------------
' SEND EMAIL notification to the requester when the transfer has been updated in BASE
'-----------------------------------------------------------------------------

sub sendTransferBaseUpdatedEmail		
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
				
	emailTo = "nicole.aquilina@silklogistics.com.au"
	emailCc = "neil.jones@silklogistics.com.au"
	emailBcc = "logistics-aus@music.yamaha.com"
	'emailBcc = "Harsono_Setiono@gmx.yamaha.com"
		
	emailFrom 		= "automailer@music.yamaha.com"		
	emailSubject 	= "Your Transfer Request (" & session("warehouse") & ") has been updated in BASE"
		
	emailBodyText   = "Hi Nicole," & vbCrLf _	
					& "" & vbCrLf _
					& "The transfer request from " & session("warehouse") & " created at " & session("date_created") & " has been updated in BASE." & vbCrLf _
					& "" & vbCrLf _
					& "Thank you." & vbCrLf _
					& "" & vbCrLf _
					& "Yamaha Logistics Division" & vbCrLf _	
					& "" & vbCrLf _					
					& "This is an automated email - please do not reply to this email."
		
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
		
	strMessageText = "An email has been sent to Nicole (3K) informing that the transfer has been updated in BASE."
end sub

sub main
	call UTL_validateLogin  
	call sendTransferBaseUpdatedEmail
end sub

call main

dim strMessageText
%>
</head>
<body>
<table width="100%" cellpadding="0" cellspacing="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><p><%= strMessageText %></p>
      <p>Click <a href="list_transfer.asp">here</a> to go back to Transfer Requests.</p></td>
  </tr>
</table>
</body>
</html>