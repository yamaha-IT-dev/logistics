<%
'----------------------------------------------------------------------------------------
' UPDATE INVENTORY
'----------------------------------------------------------------------------------------
Function updateInventory(intID, strDepartment, strType, strProduct, strTariffCode, strTariffRate, strDuty, intFTA, strUsername)
	dim strSQL

	Call OpenWorkflowDataBase()

	strSQL = "UPDATE workflow_inventory_master_reference SET "
	strSQL = strSQL & "tariff_code = '" & Server.HTMLEncode(strTariffCode) & "',"
	strSQL = strSQL & "tariff_rate = '" & Server.HTMLEncode(strTariffRate) & "',"
	strSQL = strSQL & "duty = '" & Server.HTMLEncode(strDuty) & "',"
	strSQL = strSQL & "fta = '" & Server.HTMLEncode(intFTA) & "',"
	strSQL = strSQL & "last_update_date = GetDate(),"
	strSQL = strSQL & "last_update_people = '" & strUsername & "' WHERE id = " & intID

	'response.Write strSQL	
	on error resume next
	conn.Execute strSQL

	On error Goto 0

	if err <> 0 then
		strMessageText = err.description
	else
		call sendInventoryApprovalEmail(intID, strDepartment, strType, strProduct)
	end if

	Call CloseDataBase()
end function

'----------------------------------------------------------------------------------------
' SEND INVENTORY GM APPROVAL NOTIFICATION
'----------------------------------------------------------------------------------------
Function sendInventoryApprovalEmail(intID, strDepartment, strType, strProduct)
	dim strFirstname
	
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
		
	select case	 strDepartment 
		case "AV"
			emailTo	= "harsono_setiono@gmx.yamaha.com"
			emailCc	= "gandi_gan@gmx.yamaha.com"
			strFirstname = "Simon"
		case "MPD"
			emailTo	= "harsono_setiono@gmx.yamaha.com"
			emailCc	= "gandi_gan@gmx.yamaha.com"
			strFirstname = "Mark"
	end select
	
	emailTo 	= "Harsono_Setiono@gmx.yamaha.com"
	emailSubject = "Inventory Master Maintenance - Needs your approval"
				
	emailBodyText =	"G'day " & strFirstname & "," & vbCrLf _
					& " " & vbCrLf _
					& "This following Inventory Master Maintenance needs your approval." & vbCrLf _
					& " " & vbCrLf _
					& "Type    : " & strType & vbCrLf _
					& "Product : " & strProduct & vbCrLf _
					& " " & vbCrLf _
					& "Please click on the below link to approve it:" & vbCrLf _
					& "http://intranet:89/workflow/Inventory_Master/Default.aspx?id=3" & intID & vbCrLf _
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
	
	Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
end function
%>