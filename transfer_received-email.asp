<!--#include file="include/connection_it.asp " -->
<% strSection = "freight" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Freight - Send Received Transfer Notification</title>
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/validations.js"></script>
<%
'-----------------------------------------------------------------------------
' SEND EMAIL notification to the Yamaha when the freight has been received
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
				
	'emailTo = "johanna_scholes@gmx.yamaha.com"
	emailTo = "logistics-aus@music.yamaha.com"
	'emailBcc = "harsono_setiono@gmx.yamaha.com"
		
	emailFrom 		= "automailer@music.yamaha.com"
	emailSubject 	= "Transfer Request (" & session("warehouse") & ") has been received"
		
	emailBodyText   = "Hello," & vbCrLf _
					& "" & vbCrLf _
					& "The transfer request from " & session("warehouse") & " created at " & session("date_created") & " has been received." & vbCrLf _
					& "" & vbCrLf _
					& "Thank you." & vbCrLf _
					& "" & vbCrLf _
					& "" & session("UsrUserName") & " " & vbCrLf _
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
		
	strMessageText = "An email has been sent to Yamaha Logistics informing that the transfer request has been received."
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
      <p>Click <a href="list_transfer.asp">here</a> to go back to Transfer Requests.</p></td>
  </tr>
</table>
</body>
</html>