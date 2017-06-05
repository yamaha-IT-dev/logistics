<%
function checkUsername(strEmail)
	dim strSQL
	dim rs	
		
	call OpenDataBase()
	
	strSQL = "SELECT * FROM tbl_users WHERE email = '" & strEmail & "' "
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.Open strSQL, conn
	
	if rs.EOF then
		strMessageText = "<div class=""alert alert-danger""><img src=""../images/icon_cross.png""> Email not found. Please retry.</div>"
    else
		dim strFirstname, strUsername, strPassword
		
		strFirstname	= rs("firstname")
		strUsername		= rs("username")
    	strPassword 	= rs("password")
		
		dim oMail, iConf, Flds
		
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
				
		emailFrom 	  	= "automailer@music.yamaha.com"
		emailTo 	  	= strEmail
		emailSubject  	= "Logistics Portal Login"
		
		emailBodyText =	 "G'day " & strFirstname & "!" & vbCrLf _
						& "" & vbCrLf _
						& "As requested." & vbCrLf _
						& "" & vbCrLf _
						& "U: " & strUsername & vbCrLf _
						& "P: " & strPassword & vbCrLf _
						& "" & vbCrLf _
						& "This is an automated email. Please do not reply to this email."
				
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
		
		strMessageText = "<div class=""alert alert-success""><img src=""../images/icon_check.png""> Please check your inbox.</div>"
	end if	
	
	call CloseDataBase()
end function
%>