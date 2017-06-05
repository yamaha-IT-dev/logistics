<%
'-----------------------------------------------
' LIST SPECIFIC LOG
'-----------------------------------------------
function listLog(logType, logForeignKey)
    dim strSQL
	dim intRecordCount
	
    call OpenDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic			
	rs.PageSize = 200
	
	strSQL = "SELECT * FROM tbl_log_logistics WHERE logForeignKey = '" & logForeignKey & "' AND logType = '" & logType & "' ORDER BY logDate"
	
	rs.Open strSQL, conn
	
	intRecordCount = rs.recordcount	

    strLogList = ""
	
	if not DB_RecSetIsEmpty(rs) Then
	
		For intRecord = 1 To rs.PageSize
		
			strLogList = strLogList & "<tr><td width=""20%"">" & trim(rs("logUsername")) & "</td>"
			strLogList = strLogList & "<td width=""25%"">" & trim(rs("logActivity")) & "</td>"
			strLogList = strLogList & "<td width=""55%"">" & WeekDayName(WeekDay(rs("logDate"))) & ", " & FormatDateTime(rs("logDate"),1) & " at " & FormatDateTime(rs("logDate"),3) & "</td></tr>"
			
			rs.movenext
			
			If rs.EOF Then Exit For
		next
	else
        strLogList = "<tr><td colspan=""3"">&nbsp;</td></tr>"
	end if
	
	strLogList = strLogList & "<tr>"
	
    call CloseDataBase()
end function

'-----------------------------------------------
' ADD LOG (NEW RECORD)
'-----------------------------------------------
function addLogNew(logType, logForeignKey, logUsername, logActivity)
	dim strSQL

	Call OpenDataBase()
		
	strSQL = "INSERT INTO tbl_log_logistics (logType, logForeignKey, logUsername, logActivity) VALUES ("
	strSQL = strSQL & " '" & logType & "',"
	strSQL = strSQL & " '" & logForeignKey & "',"
	strSQL = strSQL & " '" & Trim(logUsername) & "',"
	strSQL = strSQL & " '" & Trim(logActivity) & "')"

	on error resume next
	conn.Execute strSQL
	
	'response.Write strSQL
	
	if err <> 0 then
		strMessageText = err.description
	'else
	'	strMessageText = "This activity has been logged."
	end if
	
	Call CloseDataBase()
end function

'-----------------------------------------------
' ADD LOG (UPDATE RECORD)
'-----------------------------------------------
function addLog(logType, logForeignKey, logUsername, logActivity)
	dim strSQL
 
	Call OpenDataBase()
	
	strSQL = "INSERT INTO tbl_log_logistics (logType, logForeignKey, logUsername, logActivity) VALUES ("
	strSQL = strSQL & " '" & logType & "',"
	strSQL = strSQL & " '" & logForeignKey & "',"
	strSQL = strSQL & " '" & Trim(logUsername) & "',"
	strSQL = strSQL & " '" & Trim(logActivity) & "')"

	on error resume next
	conn.Execute strSQL
	
	'response.Write strSQL
	
	if err <> 0 then
		strMessageText = err.description
	'else
	'	strMessageText = "This activity has been logged."
	end if
	
	Call CloseDataBase()
end function
%>