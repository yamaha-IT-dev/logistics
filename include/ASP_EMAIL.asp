<%
'H1**********************************************************************************
'Copyright			:Digital Zoo Pty Ltd - 2003
'Library			:ASPEML_asp_email
'Creation Date		:6 July 2003
'Version			:1.0
'Author(s)			:Simon Abriani
'
'Library Purpose	:Send a email(s) using Persits ASPEmail
'Procedures			:ASPEML_Send()
'Special Info		:NA
'Revision History	:################################################################
'H2**********************************************************************************
'-----------------------------------
'EMAIL PRIORITY
'-----------------------------------
const ASPEML_HIGH_PRIORITY	= 1
const ASPEML_MED_PRIORITY	= 3
const ASPEML_LOW_PRIORITY	= 5

'P1********************************************************************************
'Procedure		:ASPEML_send
'Arguments		:strHosts as string - mail server(s)
'				:strPort as integer
'				:strFromAddress as string
'				:strFromName as string
'				:strSubject as string
'				:strBody as string
'				:blnIsHtml as boolean
'				:intPriority as integer
'				:strEmails as string - delimitied list of emails
'				:strCCEmails as string - delimited list of emails to cc
'				:strBccEmails as string - delimited list of emails to bcc
'				:strAttachments as string - delimited list of file locations to attach
'				:strDelim as string - delimiter
'				:strErrMsg as string - error message variable.
'Purpose		:Sends an email using ASPEmail
'Output			:True if successful, otherwise false.
'Dependancies	:Persits.MailSender
'Special Info	:NA
'P2********************************************************************************
function ASPEML_send(strHosts, strPort, strFromAddress, strFromName, strSubject, strBody, blnIsHtml, intPriority, strEmails, strCCEmails, strBccEmails, strAttachments, strDelim, strErrMsg)
	Dim oMail
	Dim arrEmails
	dim arrCCEmails
	dim arrBccEmails
	dim arrAttachments
	Dim intI
	
	If 1=1 Then
		On Error Resume Next
				
		Set oMail = Server.CreateObject("Persits.MailSender")

		oMail.Host		= strHosts
		oMail.Port		= strPort
		oMail.From		= strFromAddress
		oMail.FromName	= strFromName
		oMail.Priority	= intPriority
		oMail.IsHTML	= blnIsHtml
		oMail.Subject	= strSubject
		oMail.Body		= strBody
		

		if len(strEmails) > 0 then
			arrEmails = split(strEmails, strDelim)
			for intI = lbound(arrEmails) to ubound(arrEmails)
			    'For all test environment, only certain email can be used to send
                if ENVIRONMENT = "LIVE" then
                    oMail.AddAddress arrEmails(intI)
                else
                    'only certain email can be sent
                    if UTL_FindListItem(ST_VALID_TEST_EMAIL_ADDR, arrEmails(intI), ",") >= 0 then
                        oMail.AddAddress arrEmails(intI)
                    end if
                end if
			    
			next
		end if

		if len(strCCEmails) > 0 then
			arrCCEmails = split(strCCEmails, strDelim)
			for intI = lbound(arrCCEmails) to ubound(arrCCEmails)
				'For all test environment, only certain email can be used to send
                if ENVIRONMENT = "LIVE" then
                    oMail.AddCC arrCCEmails(intI)
                else
                    'only certain email can be sent
                    if UTL_FindListItem(ST_VALID_TEST_EMAIL_ADDR, arrCCEmails(intI), ",") >= 0 then
                        oMail.AddCC arrCCEmails(intI)
                    end if
                end if
			next
		end if

		if len(strBccEmails) > 0 then
			arrBccEmails = split(strBccEmails, strDelim)
			for intI = lbound(arrBccEmails) to ubound(arrBccEmails)
				'For all test environment, only certain email can be used to send
                if ENVIRONMENT = "LIVE" then
                    oMail.AddBCC arrBccEmails(intI)
                else
                    'only certain email can be sent
                    if UTL_FindListItem(ST_VALID_TEST_EMAIL_ADDR, arrBccEmails(intI), ",") >= 0 then
                        oMail.AddBCC arrBccEmails(intI)
                    end if
                end if
			next
		end if


        

		if len(strAttachments) > 0 then
			arrAttachments = split(strAttachments, strDelim)
			for intI = lbound(arrAttachments) to ubound(arrAttachments)
				oMail.AddAttachment arrAttachments(intI)
			next
		end if
		
		
		
		oMail.Queue = APP_EML_QUEUE		'do NOT uncomment until AspEmail Premium features are registered
		oMail.Send
		
		


			'debugging if need it
			'--------------------------------------------------------------
			'dim prevPAGE_debugState
			'prevPAGE_debugState = PAGE_debugState
			'PAGE_debugState = true

			'Call DEBUG_print(array("strDelim			-- " & strDelim, False))
			'Call DEBUG_print(array("strEmails			-- " & strEmails, False))
			'Call DEBUG_print(array("strCCEmails		-- " & strCCEmails, False))
			'Call DEBUG_print(array("strBccEmails		-- " & strBccEmails, False))
			'Call DEBUG_print(array("strAttachments	-- " & strAttachments & "<br>", False))
			
			'Call DEBUG_print(array("oMail.AltBody		-- " & oMail.AltBody, False))
			'Call DEBUG_print(array("oMail.Body		-- " & oMail.Body, False))
			'Call DEBUG_print(array("oMail.CharSet		-- " & oMail.CharSet, False))
			'Call DEBUG_print(array("oMail.ContentTransferEncoding	-- " & oMail.ContentTransferEncoding, False))
			'Call DEBUG_print(array("oMail.Expires 	-- " & oMail.Expires , False))
			'Call DEBUG_print(array("oMail.From		-- " & oMail.From, False))
			'Call DEBUG_print(array("oMail.FromName	-- " & oMail.FromName, False))
			'Call DEBUG_print(array("oMail.Helo		-- " & oMail.Helo, False))
			'Call DEBUG_print(array("oMail.Host		-- " & oMail.Host, False))
			'Call DEBUG_print(array("oMail.IsHTML		-- " & oMail.IsHTML, False))
			'Call DEBUG_print(array("oMail.MailFrom	-- " & oMail.MailFrom, False))
			'Call DEBUG_print(array("oMail.Password	-- " & oMail.Password, False))
			'Call DEBUG_print(array("oMail.Port		-- " & oMail.Port, False))
			'Call DEBUG_print(array("oMail.Priority	-- " & oMail.Priority, False))
			'Call DEBUG_print(array("oMail.Queue		-- " & oMail.Queue, False))
			'Call DEBUG_print(array("oMail.QueueFileName	-- " & oMail.QueueFileName, False))
			'Call DEBUG_print(array("oMail.Subject		-- " & oMail.Subject, False))
			'Call DEBUG_print(array("oMail.Timeout		-- " & oMail.Timeout, False))
			'Call DEBUG_print(array("oMail.Timestamp	-- " & oMail.Timestamp, False))
			'Call DEBUG_print(array("oMail.Username	-- " & oMail.Username, False))

			'Call DEBUG_print(array("email sent? = " & err.number, False))
			
			'PAGE_debugState = prevPAGE_debugState
			'--------------------------------------------------------------

		If Err <> 0 Then
			strErrMsg = strErrMsg & Err.Description & " (" & err.number & ")"
			ASPEML_send = false
		else
			ASPEML_send = true
		End If
		on error goto 0

		set oMail = nothing
	Else
		ASPEML_send = true
	End If
end function


%>