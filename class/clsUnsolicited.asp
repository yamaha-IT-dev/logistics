<%
'----------------------------------------------------------------------------------------
' ADD UNSOLICITED GOODS
'----------------------------------------------------------------------------------------
function addUnsolicited(unsDepartment, unsItemCode, unsDescription, unsConnote, unsGRA, unsDealer, unsShipmentNo, unsQty, unsInstruction, unsComments, unsCreatedBy)
	dim strSQL
			
	call OpenDataBase()
		
	strSQL = "INSERT INTO logistic_unsolicited ("
	strSQL = strSQL & "unsDepartment,"
	strSQL = strSQL & "unsItemCode,"
	strSQL = strSQL & "unsDescription,"
	strSQL = strSQL & "unsConnote,"
	strSQL = strSQL & "unsGRA,"
	strSQL = strSQL & "unsDealer,"
	strSQL = strSQL & "unsShipmentNo,"
	strSQL = strSQL & "unsQty,"
	strSQL = strSQL & "unsInstruction,"
	strSQL = strSQL & "unsComments,"
	strSQL = strSQL & "unsCreatedBy"
	strSQL = strSQL & ") VALUES ("
	strSQL = strSQL & "'" & unsDepartment & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(unsItemCode) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(unsDescription) & "',"	
	strSQL = strSQL & "'" & Server.HTMLEncode(unsConnote) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(unsGRA) & "',"	
	strSQL = strSQL & "'" & Server.HTMLEncode(unsDealer) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(unsShipmentNo) & "',"
	strSQL = strSQL & "'" & unsQty & "',"
	strSQL = strSQL & "'" & unsInstruction & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(unsComments) & "',"	
	strSQL = strSQL & "'" & unsCreatedBy & "')"
	'strSQL = strSQL & " CONVERT(datetime,'" & strDateReceived & "',103),"	
	
	'response.Write strSQL
	on error resume next
	conn.Execute strSQL
	
	if err <> 0 then
		strMessageText = err.description
	else	
		Response.Redirect("list_unsolicited.asp")
	end if

	call CloseDataBase()
end function

'----------------------------------------------------------------------------------------
' GET UNSOLICITED GOODS
'----------------------------------------------------------------------------------------
function getUnsolicited(unsID)
	dim strSQL

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strSQL = "SELECT * FROM logistic_unsolicited WHERE unsID = " & unsID

	rs.Open strSQL, conn

    if not DB_RecSetIsEmpty(rs) Then
		session("unsDepartment") 	= rs("unsDepartment")
		session("unsItemCode") 		= rs("unsItemCode")
		session("unsDescription") 	= rs("unsDescription")
		session("unsConnote")		= rs("unsConnote")
		session("unsGRA") 			= rs("unsGRA")
		session("unsDealer") 		= rs("unsDealer")
		session("unsShipmentNo") 	= rs("unsShipmentNo")
		session("unsQty") 			= rs("unsQty")		
		session("unsInstruction") 	= rs("unsInstruction")		
		session("unsComments") 		= rs("unsComments")
		session("unsDateCreated")	= rs("unsDateCreated")
		session("unsCreatedBy") 	= rs("unsCreatedBy")
		session("unsDateModified")	= rs("unsDateModified")
		session("unsModifiedBy") 	= rs("unsModifiedBy")
		session("unsStatus") 		= rs("unsStatus")
    end if

    call CloseDataBase()
end function

'----------------------------------------------------------------------------------------
' UPDATE UNSOLICITED GOODS
'----------------------------------------------------------------------------------------
function updateUnsolicited(unsID, unsDepartment, unsItemCode, unsDescription, unsConnote, unsGRA, unsDealer, unsShipmentNo, unsQty, unsInstruction, unsComments, unsStatus, unsModifiedBy)
	dim strSQL

	call OpenDataBase()

	strSQL = "UPDATE logistic_unsolicited SET "
	strSQL = strSQL & "unsDepartment = '" & unsDepartment & "',"
	strSQL = strSQL & "unsItemCode = '" & Server.HTMLEncode(unsItemCode) & "',"
	strSQL = strSQL & "unsDescription = '" & Server.HTMLEncode(unsDescription) & "',"
	strSQL = strSQL & "unsConnote = '" & Server.HTMLEncode(unsConnote) & "',"
	strSQL = strSQL & "unsGRA = '" & Server.HTMLEncode(unsGRA) & "',"
	strSQL = strSQL & "unsDealer = '" & Server.HTMLEncode(unsDealer) & "',"
	strSQL = strSQL & "unsShipmentNo = '" & Server.HTMLEncode(unsShipmentNo) & "',"
	strSQL = strSQL & "unsQty = '" & unsQty & "',"
	strSQL = strSQL & "unsInstruction = '" & unsInstruction & "',"
	strSQL = strSQL & "unsComments = '" & Server.HTMLEncode(unsComments) & "',"
	strSQL = strSQL & "unsDateModified = GetDate(),"
	strSQL = strSQL & "unsModifiedBy = '" & unsModifiedBy & "',"
	strSQL = strSQL & "unsStatus = '" & unsStatus & "' WHERE unsID = " & unsID

	'response.Write strSQL

	on error resume next
	conn.Execute strSQL

	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "The record has been updated."
	end if

	call CloseDataBase()
end function
%>