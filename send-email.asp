<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Freight - Send Transfer Pickup Email Notification</title>
<%
'-----------------------------------------------------------------------------
' SEND EMAIL notification to the requester when the transfer has been picked up
'-----------------------------------------------------------------------------

sub sendTransferPickupEmail

		Set JMail=CreateObject("JMail.SMTPMail")
		
		JMail.ServerAddress = "smtp.bne.server-mail.com"
		JMail.Subject		= "Transfer Request has been picked up"
		JMail.Sender		= "au_webmaster@gmx.yamaha.com"
		JMail.SenderName	= "Yamaha Logistics"
		
		JMail.AddRecipient (Trim(Request("email")))
				
		JMail.Body    	= "Hi there," & vbCrLf _
						& "" & vbCrLf _
						& "Your transfer request created at " & Trim(Request("created")) & " has been picked up." & vbCrLf _ 
						& "" & vbCrLf _
						& "Thank you." & vbCrLf _
						& ""  & vbCrLf _
						& "Yamaha Logistics Division" & vbCrLf _	
						& ""  & vbCrLf _					
						& "This is an automated email - please do not reply to this email."
				
		'JMail.BodyFormat = 0
		'JMail.MailFormat = 0
		JMail.Execute
		
		set JMail=nothing	
		
		strMessageText = "An email has been sent to notify that the transfer has been picked up."
end sub

sub main
	call sendTransferPickupEmail
end sub

call main

dim strMessageText
%>
</head>
<body>
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td class="first_content"><p><%= strMessageText %></p>
      <p>Click <a href="http://intranet/logistics/list_transfer.asp">here</a> to go back to Transfer Requests.</p></td>
  </tr>
</table>
</body>
</html>