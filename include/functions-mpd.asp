<%
'-----------------------------------------------
'				GET DEALER DETAILS
'-----------------------------------------------
function getDealerDetails
    dim strSqlQuery
    dim rs

    call OpenDataBase()

    intDealerId = request("dealer_id")
	strSqlQuery = "EXEC spGetDealer " & intDealerId
			
	set rs = server.CreateObject("ADODB.Recordset")
    set rs = conn.execute(strSqlQuery)
           
    if not DB_RecSetIsEmpty(rs) Then
		session("dealer_name") = rs("dealer_name")
		session("rebate_both") = rs("rebate_both")
		session("rebate_trad") = rs("rebate_trad")
		session("rebate_pro") = rs("rebate_pro")
		session("BOY_roadshow9") = rs("BOY_roadshow9")
		session("BOY_people") = rs("BOY_people")
		session("F06") = rs("F06")
		session("F07") = rs("F07")
		session("F08") = rs("F08")
		session("F09") = rs("F09")
		session("address") = rs("address")
		session("po_box") = rs("po_box")
		session("suburb") = rs("suburb")
		session("state") = rs("state")
		session("postcode") = rs("postcode")
		session("sales_manager") = rs("sales_manager")
		session("phone") = rs("phone")
		session("fax") = rs("fax")
		session("principal") = rs("principal")
		session("principal_email") = rs("principal_email")	
		session("principal_mobile") = rs("principal_mobile")
		session("store_manager") = rs("store_manager")
		session("store_manager_email") = rs("store_manager_email")
		session("store_manager_mobile") = rs("store_manager_mobile")
		session("website") = rs("website")
		session("email_sales") = rs("email_sales")
		session("email_legal") = rs("email_legal")
		session("staff1") = rs("staff1")
		session("staff1_email") = rs("staff1_email")
		session("staff2") = rs("staff2")
		session("staff2_email") = rs("staff2_email")
		session("staff3") = rs("staff3")
		session("staff3_email") = rs("staff3_email")
		session("staff4") = rs("staff4")
		session("staff4_email") = rs("staff4_email")
		session("staff5") = rs("staff5")
		session("staff5_email") = rs("staff5_email")	
		session("piano_cont") = rs("piano_cont")
		session("piano_GS") = rs("piano_GS")
		session("piano_CS") = rs("piano_CS")
		session("piano_premium") = rs("piano_premium")
		session("disklavier") = rs("disklavier")
		session("clavinova_std") = rs("clavinova_std")
		session("clavinova_plt") = rs("clavinova_plt")
		session("modus") = rs("modus")
		session("digital_piano_openline") = rs("digital_piano_openline")
		session("portable_keyboard_openline") = rs("portable_keyboard_openline")
		session("portable_keyboard_std") = rs("portable_keyboard_std")
		session("portable_keyboard_pro") = rs("portable_keyboard_pro")
		session("band_std") = rs("band_std")
		session("band_PN") = rs("band_PN")
		session("string") = rs("string")
		session("CP") = rs("CP")
		session("guitar_openline") = rs("guitar_openline")
		session("guitar_ranging") = rs("guitar_ranging")
		session("handcrafted_guitars") = rs("handcrafted_guitars")
		session("vox_openline") = rs("vox_openline")
		session("vox_ranging") = rs("vox_ranging")
		session("vox_guitars") = rs("vox_guitars")
		session("acoustic_drums_openline") = rs("acoustic_drums_openline")
		session("acoustic_drum_dealer") = rs("acoustic_drum_dealer")
		session("drums_ranging") = rs("drums_ranging")
		session("drums_DTX_openline") = rs("drums_DTX_openline")
		session("drums_DTX_ranging") = rs("drums_DTX_ranging")
		session("drums_DTXtreme") = rs("drums_DTXtreme")
		session("drums_oak_recording_custom") = rs("drums_oak_recording_custom")
		session("drum_system") = rs("drum_system")
		session("paiste_openline") = rs("paiste_openline")
		session("paiste_tree") = rs("paiste_tree")
		session("paiste_inter") = rs("paiste_inter")
		session("paiste_wall") = rs("paiste_wall")
		session("MP_openline") = rs("MP_openline")
		session("MP_ranging") = rs("MP_ranging")
		session("PA_openline") = rs("PA_openline")
		session("PA_ranging") = rs("PA_ranging")
		session("steinberg_openline") = rs("steinberg_openline")
		session("steinberg_project") = rs("steinberg_project")
		session("steinberg_cubase") = rs("steinberg_cubase")
		session("steinberg_cubase_education") = rs("steinberg_cubase_education")
		session("steinberg_nuendo") = rs("steinberg_nuendo")
		session("software_computer_vendor") = rs("software_computer_vendor")
		session("O1V") = rs("O1V")
		session("LS9") = rs("LS9")
		session("CA") = rs("CA")	
		session("comments") = rs("comments")
    	session("sales_manager_id") = rs("sales_manager_id")
		
		if rs("status") then
			session("status") 	= "1"
		else
			session("status") 	= "0"
		end if
		
    end if   	
	
    rs.Close
    set rs = nothing
    
    call CloseDataBase()
end function
'-----------------------------------------------
'				GET NEW DEALER DETAILS
'-----------------------------------------------
function getNewDealerDetails
    dim strSqlQuery
    dim rs

    call OpenDataBase()

    intDealerId = request("dealer_id")
	strSqlQuery = "EXEC spGetDealer " & intDealerId
			
	set rs = server.CreateObject("ADODB.Recordset")
    set rs = conn.execute(strSqlQuery)
           
    if not DB_RecSetIsEmpty(rs) Then
		session("dealer_name") = rs("dealer_name")		
		session("address") = rs("address")
		session("suburb") = rs("suburb")
		session("state") = rs("state")
		session("postcode") = rs("postcode")
		session("phone") = rs("phone")
		session("fax") = rs("fax")		
		session("email_sales") = rs("email_sales")		
		session("piano_cont") = rs("piano_cont")
		session("piano_GS") = rs("piano_GS")
		session("piano_CS") = rs("piano_CS")
		session("piano_premium") = rs("piano_premium")
		session("disklavier") = rs("disklavier")
		session("clavinova_std") = rs("clavinova_std")
		session("clavinova_plt") = rs("clavinova_plt")
		session("modus") = rs("modus")
		session("digital_piano_openline") = rs("digital_piano_openline")
		session("portable_keyboard_openline") = rs("portable_keyboard_openline")
		session("portable_keyboard_std") = rs("portable_keyboard_std")
		session("portable_keyboard_pro") = rs("portable_keyboard_pro")
		session("band_std") = rs("band_std")
		session("band_PN") = rs("band_PN")
		session("string") = rs("string")
		session("CP") = rs("CP")
		session("guitar_openline") = rs("guitar_openline")
		session("guitar_ranging") = rs("guitar_ranging")
		session("handcrafted_guitars") = rs("handcrafted_guitars")
		session("vox_openline") = rs("vox_openline")
		session("vox_ranging") = rs("vox_ranging")
		session("vox_guitars") = rs("vox_guitars")
		session("acoustic_drums_openline") = rs("acoustic_drums_openline")
		session("acoustic_drum_dealer") = rs("acoustic_drum_dealer")
		session("drums_ranging") = rs("drums_ranging")
		session("drums_DTX_openline") = rs("drums_DTX_openline")
		session("drums_DTX_ranging") = rs("drums_DTX_ranging")
		session("drums_DTXtreme") = rs("drums_DTXtreme")
		session("drums_oak_recording_custom") = rs("drums_oak_recording_custom")
		session("drum_system") = rs("drum_system")
		session("paiste_openline") = rs("paiste_openline")
		session("paiste_tree") = rs("paiste_tree")
		session("paiste_inter") = rs("paiste_inter")
		session("paiste_wall") = rs("paiste_wall")
		session("MP_openline") = rs("MP_openline")
		session("MP_ranging") = rs("MP_ranging")
		session("PA_openline") = rs("PA_openline")
		session("PA_ranging") = rs("PA_ranging")
		session("steinberg_openline") = rs("steinberg_openline")
		session("steinberg_project") = rs("steinberg_project")
		session("steinberg_cubase") = rs("steinberg_cubase")
		session("steinberg_cubase_education") = rs("steinberg_cubase_education")
		session("steinberg_nuendo") = rs("steinberg_nuendo")
		session("software_computer_vendor") = rs("software_computer_vendor")
		session("O1V") = rs("O1V")
		session("LS9") = rs("LS9")
		session("CA") = rs("CA")	
		session("comments") = rs("comments")		
		session("date_created") = rs("date_created")
		session("requested_by") = rs("requested_by")
		session("nsm_approval") = rs("nsm_approval")
		session("nsm_approval_date") = rs("nsm_approval_date")
		session("gm_approval") = rs("gm_approval")
		session("gm_approval_date") = rs("gm_approval_date")		
		session("credit_approval") = rs("credit_approval")
		session("credit_approval_date") = rs("credit_approval_date")		
		session("nearest_dealer") = rs("nearest_dealer")
		session("nearest_dealer_km") = rs("nearest_dealer_km")
		session("total_purchases_to_date") = rs("total_purchases_to_date")
		session("current_dealer_impact") = rs("current_dealer_impact")
		session("strategy") = rs("strategy")
		session("estimated_monthly_sales_YMA") = rs("estimated_monthly_sales_YMA")
		session("YMA_percentage") = rs("YMA_percentage")
		session("estimated_monthly_sales_total") = rs("estimated_monthly_sales_total")
		session("brand1") = rs("brand1")
		session("brand2") = rs("brand2")
		session("brand3") = rs("brand3")
		session("brand4") = rs("brand4")
		session("supplier_percentage") = rs("supplier_percentage")
		session("new_store") = rs("new_store")
		session("area_size") = rs("area_size")
		session("estimated_growth") = rs("estimated_growth")				
		session("requester_firstname") = rs("firstname")	
		session("requester_lastname") = rs("lastname")	
		session("requester_email") = rs("email")
		
		session("base_dealer_code") = rs("base_dealer_code")
		session("base_name") = rs("base_name")
		session("credit_limit") = rs("credit_limit")		
		
		if rs("status") then
			session("status") 	= "1"
		else
			session("status") 	= "0"
		end if
		
    end if   	
	
    rs.Close
    set rs = nothing
    
    call CloseDataBase()
end function

'-----------------------------------------------
'				GET REQUEST DETAILS
'-----------------------------------------------
function getRequestDetails
    dim strSqlQuery
    dim rs

    call OpenDataBase()

    ' We get the values that the user is entering
    intRequestId = request("request_id")
	strSqlQuery = "EXEC spGetRequest " & intRequestId
		
	set rs = server.CreateObject("ADODB.Recordset")
    set rs = conn.execute(strSqlQuery)
           
    if not DB_RecSetIsEmpty(rs) Then
		session("R_dealer_id") = rs("dealer_id")		
		session("R_piano_cont") = rs("piano_cont")
		session("R_piano_GS") = rs("piano_GS")
		session("R_piano_CS") = rs("piano_CS")
		session("R_piano_premium") = rs("piano_premium")
		session("R_disklavier") = rs("disklavier")
		session("R_clavinova_std") = rs("clavinova_std")
		session("R_clavinova_plt") = rs("clavinova_plt")
		session("R_modus") = rs("modus")
		session("R_digital_piano_openline") = rs("digital_piano_openline")
		session("R_portable_keyboard_openline") = rs("portable_keyboard_openline")
		session("R_portable_keyboard_std") = rs("portable_keyboard_std")
		session("R_portable_keyboard_pro") = rs("portable_keyboard_pro")
		session("R_band_std") = rs("band_std")
		session("R_band_PN") = rs("band_PN")
		session("R_string") = rs("string")
		session("R_CP") = rs("CP")
		session("R_guitar_openline") = rs("guitar_openline")
		session("R_guitar_ranging") = rs("guitar_ranging")
		session("R_handcrafted_guitars") = rs("handcrafted_guitars")
		session("R_vox_openline") = rs("vox_openline")
		session("R_vox_ranging") = rs("vox_ranging")
		session("R_vox_guitars") = rs("vox_guitars")
		session("R_acoustic_drums_openline") = rs("acoustic_drums_openline")
		session("R_acoustic_drum_dealer") = rs("acoustic_drum_dealer")
		session("R_drums_ranging") = rs("drums_ranging")
		session("R_drums_DTX_openline") = rs("drums_DTX_openline")
		session("R_drums_DTX_ranging") = rs("drums_DTX_ranging")
		session("R_drums_DTXtreme") = rs("drums_DTXtreme")
		session("R_drums_oak_recording_custom") = rs("drums_oak_recording_custom")
		session("R_drum_system") = rs("drum_system")
		session("R_paiste_openline") = rs("paiste_openline")
		session("R_paiste_tree") = rs("paiste_tree")
		session("R_paiste_inter") = rs("paiste_inter")
		session("R_paiste_wall") = rs("paiste_wall")
		session("R_MP_openline") = rs("MP_openline")
		session("R_MP_ranging") = rs("MP_ranging")
		session("R_PA_openline") = rs("PA_openline")
		session("R_PA_ranging") = rs("PA_ranging")
		session("R_steinberg_openline") = rs("steinberg_openline")
		session("R_steinberg_project") = rs("steinberg_project")
		session("R_steinberg_cubase") = rs("steinberg_cubase")
		session("R_steinberg_cubase_education") = rs("steinberg_cubase_education")
		session("R_steinberg_nuendo") = rs("steinberg_nuendo")
		session("R_software_computer_vendor") = rs("software_computer_vendor")
		session("R_O1V") = rs("O1V")
		session("R_LS9") = rs("LS9")
		session("R_CA") = rs("CA")	
		session("R_comments") = rs("comments")
    	'session("R_sales_manager_id") = rs("sales_manager_id")
		
		session("R_last_modified_date") = rs("last_modified_date")
		session("R_last_modified_by") = rs("last_modified_by")
		session("R_date_submitted") = rs("date_submitted")
		session("R_requested_by") = rs("requested_by")
		session("R_nsm_approval") = rs("nsm_approval")
		session("R_nsm_approval_date") = rs("nsm_approval_date")
		session("R_gm_approval") = rs("gm_approval")
		session("R_gm_approval_date") = rs("gm_approval_date")
		
		session("R_nearest_dealer") = rs("nearest_dealer")
		session("R_nearest_dealer_km") = rs("nearest_dealer_km")
		session("R_total_purchases_to_date") = rs("total_purchases_to_date")
		session("R_current_dealer_impact") = rs("current_dealer_impact")
		session("R_strategy") = rs("strategy")
		session("R_estimated_monthly_sales_YMA") = rs("estimated_monthly_sales_YMA")
		session("R_YMA_percentage") = rs("YMA_percentage")
		session("R_estimated_monthly_sales_total") = rs("estimated_monthly_sales_total")
		session("R_brand1") = rs("brand1")
		session("R_brand2") = rs("brand2")
		session("R_brand3") = rs("brand3")
		session("R_brand4") = rs("brand4")
		session("R_supplier_percentage") = rs("supplier_percentage")
		session("R_general_comments") = rs("general_comments")
		session("R_maximiser_upload") = rs("maximiser_upload")		
		
		session("line_requester_firstname") = rs("firstname")	
		session("line_requester_lastname") = rs("lastname")	
		session("line_requester_email") = rs("email")	
		
		if rs("status") then
			session("status") 	= "1"
		else
			session("status") 	= "0"
		end if
		
    end if   	
	
    rs.Close
    set rs = nothing
    
    call CloseDataBase()
end function

'-----------------------------------------------
'				GET STATE
'-----------------------------------------------
function getState
    dim arrStateFillText
    dim arrStateFillID
    dim intCounter

    arrStateFillText        = split(arrStateText, ",")
    arrStateFillID 		    = split(arrStateID, ",")
    
    strStateList = strStateList & "<option value='0'>...</option>"
    
    ' We check if there is anything
    if isarray(arrStateFillID) then
        if ubound(arrStateFillID) > 0 then
        
            for intCounter = 0 to ubound(arrStateFillID)
                
                if trim(session("state")) = trim(arrStateFillID(intCounter)) then
                    strStateList = strStateList & "<option selected value=" & arrStateFillID(intCounter) & ">" & arrStateFillText(intCounter) & "</option>"
                else
                   	strStateList = strStateList & "<option value=" & arrStateFillID(intCounter) & ">" & arrStateFillText(intCounter) & "</option>"
                end if
             
            next
        end if
    
    end if

end function
'-----------------------------------------------
'				GET USER
'-----------------------------------------------
function getUser
    dim strSqlQuery
    dim rs

    call OpenDataBase()
    
	strSqlQuery = "EXEC spListUser"
		
	set rs = server.CreateObject("ADODB.Recordset")
    set rs = conn.execute(strSqlQuery)
    
    strUserList = strUserList & "<option value=''>...</option>"
    
    if not DB_RecSetIsEmpty(rs) Then
        do until rs.EOF
        
            if trim(session("sales_manager_id")) = trim(rs("user_id")) then
                strUserList = strUserList & "<option selected value=" & rs("user_id") & ">" & rs("firstname") & " " & rs("lastname") & "</option>"
            else
                strUserList = strUserList & "<option value=" & rs("user_id") & ">" & rs("firstname") & " " & rs("lastname") & "</option>"
            end if
                    
        rs.Movenext
        loop    
    
    end if
    
    rs.Close
    set rs = nothing
    
    call CloseDataBase()

end function

'-----------------------------------------------
'				LIST USERS
'-----------------------------------------------

function listUser
    dim strSqlQuery
    dim rs

    call OpenDataBase()
    
	strSqlQuery = "EXEC spListUser"
		
	set rs = server.CreateObject("ADODB.Recordset")
    set rs = conn.execute(strSqlQuery)
    
    strUsersList = strUsersList & "<option value=''>...</option>"
    
    if not DB_RecSetIsEmpty(rs) Then
        do until rs.EOF                    
        	strUsersList = strUsersList & "<option value=" & rs("user_id") & ">" & rs("firstname") & " " & rs("lastname") & "</option>"                               
        rs.Movenext
        loop    
    
    end if
    
    rs.Close
    set rs = nothing
    
    call CloseDataBase()

end function

'-------------------------------------------------------------
'	New Dealer APPROVAL PROCESS by National Sales Manager
'-------------------------------------------------------------
function approveNewDealerNSM	
	dim sql
	
	OpenDataBase()
	
	sql = "UPDATE tbl_dealers SET nsm_approval = 1, nsm_approval_date = getDate(), last_modified_date = getDate() WHERE dealer_id = " & session("dealer_id")
	
	on error resume next
	conn.Execute sql
	
	On error Goto 0  
	
	if err <> 0 then
		strMessageText = err.description
	else 
		strMessageText = "NSM Approved!"			
	end if 
	
	conn.close
end function

'-------------------------------------------------------------
'	New Dealer APPROVAL PROCESS by General Manager
'-------------------------------------------------------------

function approveNewDealerGM	
	dim sql
	
	OpenDataBase()
	
	sql = "UPDATE tbl_dealers SET gm_approval = 1, gm_approval_date = getDate(), last_modified_date = getDate(), status = 1 WHERE dealer_id = " & session("dealer_id")
		  
	on error resume next
	conn.Execute sql
	
	On error Goto 0  
	
	if err <> 0 then
		strMessageText = err.description
	else 
		strMessageText = "GM Approved!"			
	end if 
	
	conn.close
	
end function

'-------------------------------------------------------------
' New Dealer APPROVAL PROCESS by Credit Department
'-------------------------------------------------------------

function approveNewDealerCredit
	dim sql
	
	OpenDataBase()
	
	sql = "UPDATE tbl_dealers SET credit_approval = 1, credit_approval_date = getDate(), last_modified_date = getDate() WHERE dealer_id = " & session("dealer_id")
		  
	on error resume next
	conn.Execute sql
	
	On error Goto 0  
	
	if err <> 0 then
		strMessageText = err.description
	else 
		strMessageText = "Credit Approved!"			
	end if 
	
	conn.close
end function

'-------------------------------------------------------------
' New Product Extension APPROVAL PROCESS by National Sales Manager
'-------------------------------------------------------------

function approveNewRequestNSM	
	dim sql
	
	OpenDataBase()
	
	sql = "UPDATE tbl_requests SET nsm_approval = 1, nsm_approval_date = getDate(), last_modified_date = getDate() WHERE request_id = " & session("request_id")
	
	on error resume next
	conn.Execute sql
	
	On error Goto 0  
	
	if err <> 0 then
		strMessageText = err.description
	else 
		strMessageText = "NSM Approved!"			
	end if 
	
	conn.close
end function

'-------------------------------------------------------------
' New Product Extension APPROVAL PROCESS by General Manager
'-------------------------------------------------------------

function approveNewRequestGM	
	dim sql
	
	OpenDataBase()
	
	sql = "UPDATE tbl_requests SET gm_approval = 1, gm_approval_date = getDate(), last_modified_date = getDate() WHERE request_id = " & session("request_id")
		  
	on error resume next
	conn.Execute sql
	
	On error Goto 0  
	
	if err <> 0 then
		strMessageText = err.description
	else 
		strMessageText = "GM Approved!"			
	end if 
	
	conn.close
	
end function

'-----------------------------------------------------------------------------
' Update Dealer Credit Info
'-----------------------------------------------------------------------------
function updateDealerCredit
    Dim cmdObj, paraObj

    call OpenDataBase
	
    Set cmdObj = Server.CreateObject("ADODB.Command")
    cmdObj.ActiveConnection = conn
    cmdObj.CommandText = "spUpdateDealerCredit"
    cmdObj.CommandType = AdCmdStoredProc
    		
	session("base_dealer_code")	= request("txtBaseDealerCode") 
	session("base_name") 		= request("txtBaseName") 
	session("credit_limit") 	= request("txtCreditLimit") 	
	
	Set paraObj = cmdObj.CreateParameter(,AdInteger,AdParamInput,4, session("dealer_id"))
	cmdObj.Parameters.Append paraObj
	    						
	Set paraObj = cmdObj.CreateParameter(,AdVarChar,AdParamInput,255, session("base_dealer_code"))
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter(,AdVarChar,AdParamInput,255, session("base_name"))
	cmdObj.Parameters.Append paraObj
	Set paraObj = cmdObj.CreateParameter(,AdVarChar,AdParamInput,255, session("credit_limit"))
	cmdObj.Parameters.Append paraObj
	
	On Error Resume Next
	cmdObj.Execute
    On error Goto 0
	
    if CheckForSQLError(conn,"Update",strMessageText) = TRUE then
	    updateDealerCredit = FALSE
		'response.Write "not updated"
    else
		'response.write "updated!"
        strMessageText = session("dealer_name") & " has been updated"
		updateDealerCredit = TRUE
    end if

    Call DB_closeObject(paraObj)
    Call DB_closeObject(cmdObj)
	
    call CloseDataBase
end function

'-----------------------------------------------------------------------------
' SEND EMAIL NOTIFICATIONS (There's a dealer requesting new Product Extension)
'-----------------------------------------------------------------------------

sub sendEmailRequestNotificationNSM

	Dim objCDOSYSMail
	Set objCDOSYSMail = Server.CreateObject("CDO.Message")
	Dim objCDOSYSCnfg
	Set objCDOSYSCnfg = Server.CreateObject("CDO.Configuration")
	
	Set oMail = Server.CreateObject("CDO.Message")
	Set iConf = Server.CreateObject("CDO.Configuration")
	Set Flds = iConf.Fields
	
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 1
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "172.29.64.13"
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
	iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	iConf.Fields.Update
	
	 objCDOSYSMail.Configuration = objCDOSYSCnfg
	 emailTo = "Harsono_Setiono@gmx.yamaha.com"
	 'emailTo = "Steve_Legg@gmx.yamaha.com"
	 'emailCc = "Harsono_Setiono@gmx.yamaha.com"
	 'emailCc = "Rohan_Smith@gmx.yamaha.com"
	 'emailBcc = "setiono82@hotmail.com"
	 'emailBcc = "Anna_Bagnato@gmx.yamaha.com;Michael_Shade@gmx.yamaha.com"
	 emailFrom = "automailer@music.yamaha.com"
	 emailSubject = "New Product Extension Application "
	 emailBodyText = session("dealer_name") & " dealer has applied for product extension." & vbCrLf _
	 & "Please login to the dealer info system:" & vbCrLf _
	 & "http://intranet/dealers/" & vbCrLf _
	 & "Then see the below link for product access:" & vbCrLf _
	 & "http://intranet/dealers/update_new-dealer.asp?dealer_id=" & session("dealer_id") & ""
	 
	Set oMail.Configuration = iConf
	oMail.To 		= emailTo
	oMail.Cc		= emailCc
	oMail.Bcc		= emailBcc
	oMail.From 		= emailFrom
	oMail.Subject 		= emailSubject
	oMail.TextBody 		= emailBodyText
	oMail.Send
	
	Set iConf = Nothing
	Set Flds = Nothing

end sub
%>