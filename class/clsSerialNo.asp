<%
'-----------------------------------------------
' ADD SERIAL NO
'-----------------------------------------------
function addSerialNo(intID,intTypeID)
	dim strSQL
	
	dim strSerialNo
	strSerialNo 		= Replace(Request.Form("txtSerialNo"),"'","''")

	Call OpenDataBase()
		
	strSQL = "INSERT INTO tbl_serial_no ("
	strSQL = strSQL & "serial_type, "
	strSQL = strSQL & "serial_no, "
	strSQL = strSQL & "serial_associated_id, "
	strSQL = strSQL & "serial_created_by"
	strSQL = strSQL & ") VALUES ( "
	strSQL = strSQL & " '" & intTypeID & "',"
	strSQL = strSQL & " '" & Server.HTMLEncode(strSerialNo) & "',"	
	strSQL = strSQL & " '" & intID & "',"
	strSQL = strSQL & " '" & lcase(session("UsrUserName")) & "')"

	on error resume next
	conn.Execute strSQL
	
	response.write strSQL
	
	if err <> 0 then
		strMessageText = err.description
	else				
		strMessageText = "Serial No has been added."
	end if 
	
	Call CloseDataBase()
end function

'-----------------------------------------------
' LIST SERIAL NOs
'-----------------------------------------------
function listSerialNo(intID,intTypeID)
    dim strSQL
	dim intRecordCount
	
    call OpenDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic			
	rs.PageSize = 200
	
	strSQL = "SELECT * FROM tbl_serial_no "
	strSQL = strSQL & "	WHERE serial_associated_id = '" & intID & "' "
	strSQL = strSQL & "		AND serial_type = '" & intTypeID & "' "
	strSQL = strSQL & "	ORDER BY serial_no"
	
	rs.Open strSQL, conn
	
	intRecordCount = rs.recordcount	

    strSerialNoList = ""
	
	if not DB_RecSetIsEmpty(rs) Then	
	
		For intRecord = 1 To rs.PageSize
			strSerialNoList = strSerialNoList & "<tr>"
			strSerialNoList = strSerialNoList & "<td width=""90%"">" & trim(rs("serial_no")) & "</td>"
			strSerialNoList = strSerialNoList & "<td width=""10%""><a onclick=""return confirm('Are you sure you want to delete " & rs("serial_id") & " ?');"" href='delete_serial_no.asp?id=" & rs("serial_id") & "'><img src=""images/btn_delete.gif"" border=""0""></a></td>"
			strSerialNoList = strSerialNoList & "</tr>"
			rs.movenext
			
			If rs.EOF Then Exit For
		next
	else
        strSerialNoList = "<tr><td>&nbsp;</td></tr>"
	end if
	
	'strSerialNoList = strSerialNoList & "<tr>"
	
    call CloseDataBase()
end function
 
%>