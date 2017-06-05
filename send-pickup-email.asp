<!--#include file="include/connection_it.asp " -->
<% strSection = "freight" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Freight - Send Pickup Email Notification</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/validations.js"></script>
<%
'-----------------------------------------------------------------------------
' SEND EMAIL notification to the requester when the freight has been picked up
'-----------------------------------------------------------------------------
sub sendPickupEmail
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
						
	emailTo = trim(session("email"))
	emailCc = "logistics-aus@music.yamaha.com"
		
	emailFrom 		= "automailer@music.yamaha.com"		
	emailSubject 	= "Your Freight Request has been picked up"
		
	emailBodyText   = 	"Hi there," & vbCrLf _						
					&	"" & vbCrLf _
					&	"Your freight request (containing " & session("description") & ", " & session("description2") & " etc. ) created at " & session("date_created") & " to " & session("receiver_name") & " at " & session("receiver_address") & ", " & session("receiver_city") & " has been picked up." & vbCrLf _
					&	"" & vbCrLf _
					&	"Thank you." & vbCrLf _
					&	""  & vbCrLf _
					&	"Yamaha Logistics Division" & vbCrLf _	
					&	""  & vbCrLf _					
					&   "This is an automated email - please do not reply to this email."
		
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
		
	strMessageText = "An email has been sent to notify the requester that the freight has been picked up."
end sub

sub main
	call UTL_validateLogin  
	call sendPickupEmail
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
      <p>Click <a href="list_freights.asp">here</a> to go back to Freight List.</p></td>
  </tr>
</table>
</body>
</html>