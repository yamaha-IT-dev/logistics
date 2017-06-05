<%
'-----------------------------------------------
' EMC APPROVAL
'-----------------------------------------------
Function approveEMC(intItemID, strBaseCode, strModifiedBy)
	dim strSQL
	
	Call OpenDataBase()
	
	strSQL = "UPDATE yma_item_maintenance SET "
	strSQL = strSQL & "emc_approval = '1',"
	strSQL = strSQL & "emc_approval_date = GetDate(),"
	strSQL = strSQL & "date_modified = GetDate(),"
	strSQL = strSQL & "modified_by = '" & strModifiedBy & "'" 
	strSQL = strSQL & " WHERE item_id = " & intItemID
	
	'response.Write strSQL
	
	on error resume next
	conn.Execute strSQL
	
	if err <> 0 then
		strMessageText = err.description
	else
		Call addEmcApproval(intItemID, strModifiedBy)
		
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
						
		emailFrom 	= "automailer@music.yamaha.com"		
		'emailTo 	= "logistics-aus@music.yamaha.com"
		emailTo 	= "it-aus@music.yamaha.com"
		emailSubject = "New Item Maintenance - EMC Approved (Needs to be processed)"
				
		emailBodyText =	"G'day Logistics," & vbCrLf _
						& " " & vbCrLf _
						& "This following Item Maintenance has been EMC approved and needs to be processed." & vbCrLf _
						& " " & vbCrLf _
						& "Product: " & strBaseCode & vbCrLf _
						& " " & vbCrLf _
						& "Please click on the below link to view it:" & vbCrLf _
						& "http://intranet/logistics/update_item-maintenance.asp?id=" & intItemID & vbCrLf _
						& " " & vbCrLf _
						& "Thank you. (This is an automated email - please do not reply to this email)"
						
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
		
		strMessageText = "<div class=""notification_text""><img src=""images/icon_check.png""> The request has been approved by GM.</div>"
	end if
	
	Call CloseDataBase()
end Function

function addEmcApproval(intItemID, strApprovedBy)
	dim strSQL
	
	Call OpenDataBase()
	
	strSQL = "INSERT INTO yma_item_maintenance_approval (item_id, approval_type, approved, approved_by, approval_date) VALUES ("
	strSQL = strSQL & "'" & intItemID & "',"
	strSQL = strSQL & "'EMC Approved',"
	strSQL = strSQL & "1,"
	strSQL = strSQL & "'" & strApprovedBy & "',"
	strSQL = strSQL & "GetDate())"

	'response.Write strSQL
	on error resume next
	conn.Execute strSQL
	
	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "<br>EMC Log is added."
	end if
	
	Call CloseDataBase()
end function

'-----------------------------------------------
' EMC REJECTION
'-----------------------------------------------
Function rejectEMC(intItemID, strBaseCode, strModifiedBy)
	dim strSQL
	
	Call OpenDataBase()
	
	strSQL = "UPDATE yma_item_maintenance SET "
	strSQL = strSQL & "emc_approval = '2',"
	strSQL = strSQL & "emc_approval_date = GetDate(),"
	strSQL = strSQL & "date_modified = GetDate(),"
	strSQL = strSQL & "modified_by = '" & strModifiedBy & "'" 
	strSQL = strSQL & " WHERE item_id = " & intItemID
	
	'response.Write strSQL
	
	on error resume next
	conn.Execute strSQL
	
	if err <> 0 then
		strMessageText = err.description
	else		
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
						
		emailFrom 	= "automailer@music.yamaha.com"
		emailTo 	= Trim(strRequesterEmail)
		'emailCc 	= "Carolyn_Simonds@gmx.yamaha.com"	
		'emailBcc 	= "Harsono_Setiono@gmx.yamaha.com"
		emailSubject = "Item Maintenance - EMC Rejected"
						
		emailBodyText =	"Hi there," & vbCrLf _
						& " " & vbCrLf _
						& "Your Item Maintenance Request has been EMC rejected by: " & session("UsrUserName") & vbCrLf _
						& " " & vbCrLf _
						& "Dealer:  " & session("dealer_name") & vbCrLf _
						& "Address: " & session("address") & vbCrLf _
						& "         " & session("suburb") & vbCrLf _
						& "         " & session("state") & " " & session("postcode") & vbCrLf _
						& "Phone:   " & session("phone") & vbCrLf _
						& " " & vbCrLf _
						& "Please click on the below link to view it:" & vbCrLf _
						& "http://intranet/logistics/update_item-maintenance.asp?id=" & intItemID & vbCrLf _
						& " " & vbCrLf _
						& "Thank you. (This is an automated email - please do not reply to this email)"
						
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
		
		strMessageText = "<div class=""rejection_text""><img src=""images/icon_cross.jpg""> The request has been rejected by GM.</div>"
	end if
	
	Call CloseDataBase()
end Function
%>