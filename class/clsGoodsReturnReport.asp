<%
function displayNumberFormatted(strInput)	
	if IsNull(strInput) or strInput = "" or strInput = "0"  then 
		Response.Write "0"
	else
		Response.Write FormatNumber(strInput)
	end if
end function

'----------------------------------------------------------------------------------------
' Get GRA Report record
'----------------------------------------------------------------------------------------
Function getGraReport(intID)
	dim strSQL

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	
	strSQL = "SELECT * FROM yma_gra_report WHERE report_id = " & intID

	'response.Write strSQL
	
	rs.Open strSQL, conn

    if not DB_RecSetIsEmpty(rs) Then
		session("report_gra_no") 			= trim(rs("gra_no"))
		session("report_line_no") 			= trim(rs("line_no"))
		session("report_item") 				= trim(rs("item"))
		session("report_serial_no") 		= trim(rs("serial_no"))
		session("report_dealer_code") 		= trim(rs("dealer_code"))
		session("report_warranty_code") 	= trim(rs("gra_warranty_code"))
		session("report_repair_report") 	= trim(rs("repair_report"))
		session("report_labour") 			= trim(rs("labour"))
		session("report_parts") 			= trim(rs("parts"))
		session("report_gst") 				= trim(rs("gst"))
		session("report_total") 			= trim(rs("total"))
		session("report_date_created") 		= trim(rs("date_created"))
		session("report_created_by") 		= trim(rs("created_by"))
		session("report_date_modified") 	= trim(rs("date_modified"))
		session("report_modified_by") 		= trim(rs("modified_by"))
		session("report_date_received") 	= trim(rs("date_received"))
		session("report_date_repaired") 	= trim(rs("date_repaired"))
		session("report_destination") 		= trim(rs("destination"))
		session("report_pallet_no") 		= trim(rs("pallet_no"))
		session("report_invoice_exported") 	= trim(rs("invoice_exported"))
		session("report_comments") 			= trim(rs("comments"))
		session("report_status") 			= trim(rs("status"))
    end if

    call CloseDataBase()
end Function

'-----------------------------------------------
' LIST REPORT TOTAL SUM
'-----------------------------------------------
function listReportTotal
    dim strSQL
	
	dim intDay
	
    call OpenDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic			
	'rs.PageSize = 100
	
	strSQL = "SELECT sum(labour) as total_labour, sum(parts) as total_parts, sum(gst) as total_gst, sum(total) as total_total FROM yma_gra_report "
	strSQL = strSQL & " WHERE "
	strSQL = strSQL & " 	(gra_no LIKE '%" & session("gra_report_search") & "%' "
	strSQL = strSQL & "			OR item LIKE '%" & session("gra_report_search") & "%') "
	
	if session("gra_report_month") <> "" then
		strSQL = strSQL & " AND MONTH(" & session("gra_report_date") & ") = '" & trim(session("gra_report_month")) & "' "
	end if
	
	if session("gra_report_year") <> "" then
		strSQL = strSQL & "	AND YEAR(" & session("gra_report_date") & ") = '" & trim(session("gra_report_year")) & "' "
	end if
	
	strSQL = strSQL & " 	AND status LIKE '%" & session("gra_report_status") & "%' "
	
	'response.Write strSQL
	
	rs.Open strSQL, conn
	
	intRecordCount = rs.recordcount	

    strTotal = ""
	
	if not DB_RecSetIsEmpty(rs) Then	
		if IsNull(rs("total_labour")) or rs("total_labour") = "" or rs("total_labour") = "0" then 
			strTotal = strTotal & "<h3>Total Labour: -</h3>"
		else
			strTotal = strTotal & "<h3>Total Labour: $" & FormatNumber(rs("total_labour")) & "</h3>"
		end if
		
		if IsNull(rs("total_parts")) or rs("total_parts") = "" or rs("total_parts") = "0" then 
			strTotal = strTotal & "<h3>Total Parts: -</h3>"
		else
			strTotal = strTotal & "<h3>Total Parts: $" & FormatNumber(rs("total_parts")) & "</h3>"
		end if
		
		if IsNull(rs("total_gst")) or rs("total_gst") = "" or rs("total_gst") = "0" then
			strTotal = strTotal & "<h3>Total GST: -</h3>"
		else
			strTotal = strTotal & "<h3>Total GST: $" & FormatNumber(rs("total_gst")) & "</h3>"
		end if
		
		if IsNull(rs("total_total")) or rs("total_total") = "" or rs("total_total") = "0" then
			strTotal = strTotal & "<h2>Total Cost: -</h2>"
		else
			strTotal = strTotal & "<h2>Total Cost: $<u>" & FormatNumber(rs("total_total")) & "</u></h2>"
		end if
	else
        strTotal = "-"
	end if
	
    call CloseDataBase()
end function

'-----------------------------------------------
' LIST WRITEOFFS REPORT TOTAL SUM
'-----------------------------------------------
function listWriteoffsReportTotal
    dim strSQL
	
	dim intDay
	
    call OpenDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic			
	'rs.PageSize = 100
	
	strSQL = "SELECT sum(labour) as total_labour, sum(parts) as total_parts, sum(gst) as total_gst, sum(total) as total_total, sum(LIC) as total_lic FROM yma_gra_report "
	strSQL = strSQL & " LEFT JOIN "
	strSQL = strSQL & "		OPENQUERY(AS400, 'SELECT E2SOSC, E2NGTY, E2NGTM, (E2ihtn+(e2ihtn*e2kzrt/100)+(e2ihtn*e2skkr/100)) as LIC FROM EF2SP') "
	strSQL = strSQL & "				ON item = E2SOSC and E2NGTY = YEAR(getdate()) AND E2NGTM = MONTH(getdate()) "
	strSQL = strSQL & " WHERE pallet_no LIKE '%" & session("gra_report_pallet") & "%' "

	if session("gra_report_year") <> "" then
		strSQL = strSQL & " AND YEAR(date_created) = '" & trim(session("gra_report_year")) & "' "	
	end if	
	
	if session("gra_report_month") <> "" then
		strSQL = strSQL & " AND MONTH(date_created) = '" & trim(session("gra_report_month")) & "' "
	end if
	
	strSQL = strSQL & " 	AND (gra_no LIKE '%" & session("gra_report_search") & "%' "
	strSQL = strSQL & "			OR item LIKE '%" & session("gra_report_search") & "%') "
	strSQL = strSQL & " 	AND status LIKE '%" & session("gra_report_status") & "%' "
	strSQL = strSQL & " 	AND destination = 'Destroy'"
			
	'Response.Write "<br>" & strSQL
	
	rs.Open strSQL, conn
	
	intRecordCount = rs.recordcount	

    strTotal = ""
	
	if not DB_RecSetIsEmpty(rs) Then
		if IsNull(rs("total_lic")) or rs("total_lic") = "" or rs("total_lic") = "0" then
			strTotal = strTotal & "<h2>Total LIC: -</h2>"
		else
			strTotal = strTotal & "<h2>Total LIC: $" & FormatNumber(rs("total_lic")) & "</h2>"
		end if
		
		if IsNull(rs("total_labour")) or rs("total_labour") = "" or rs("total_labour") = "0" then 
			strTotal = strTotal & "<h3>Total Labour: -</h3>"
		else
			strTotal = strTotal & "<h3>Total Labour: $" & FormatNumber(rs("total_labour")) & "</h3>"
		end if
		
		if IsNull(rs("total_parts")) or rs("total_parts") = "" or rs("total_parts") = "0" then 
			strTotal = strTotal & "<h3>Total Parts: -</h3>"
		else
			strTotal = strTotal & "<h3>Total Parts: $" & FormatNumber(rs("total_parts")) & "</h3>"
		end if
		
		if IsNull(rs("total_gst")) or rs("total_gst") = "" or rs("total_gst") = "0" then
			strTotal = strTotal & "<h3>Total GST: -</h3>"
		else
			strTotal = strTotal & "<h3>Total GST: $" & FormatNumber(rs("total_gst")) & "</h3>"
		end if
		
		if IsNull(rs("total_total")) or rs("total_total") = "" or rs("total_total") = "0" then
			strTotal = strTotal & "<h2>Total Cost: -</h2>"
		else
			strTotal = strTotal & "<h2>Total Cost: $<u>" & FormatNumber(rs("total_total")) & "</u></h2>"
		end if
	else
        strTotal = "-"
	end if
	
    call CloseDataBase()
end function

'-----------------------------------------------
' LIST EXPORTED REPORT TOTAL SUM
'-----------------------------------------------
function listExportedReportTotal
    dim strSQL
	
	dim intDay
	
    call OpenDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic			
	'rs.PageSize = 100
			
	strSQL = "SELECT sum(labour) as total_labour, sum(parts) as total_parts, sum(gst) as total_gst, sum(total) as total_total FROM yma_gra_report "
	strSQL = strSQL & " WHERE (date_exported BETWEEN CONVERT(datetime,'" & trim(session("gra_exported_report_start")) & " 00:00:00',103)" & " "
	strSQL = strSQL & " 	AND CONVERT(datetime,'" & trim(session("gra_exported_report_end")) & " 23:59:59',103)" & ") "
	strSQL = strSQL & " 	AND (gra_no LIKE '%" & session("gra_exported_report_search") & "%' "
	strSQL = strSQL & "			OR item LIKE '%" & session("gra_exported_report_search") & "%') "
	strSQL = strSQL & " 	AND status = '0' "
	
	'Response.Write "<br>" & strSQL
	
	rs.Open strSQL, conn
	
	intRecordCount = rs.recordcount	

    strTotal = ""
	
	if not DB_RecSetIsEmpty(rs) Then	
		if IsNull(rs("total_labour")) or rs("total_labour") = "" or rs("total_labour") = "0" then 
			strTotal = strTotal & "<h3>Total Labour: -</h3>"
		else
			strTotal = strTotal & "<h3>Total Labour: $" & FormatNumber(rs("total_labour")) & "</h3>"
		end if
		
		if IsNull(rs("total_parts")) or rs("total_parts") = "" or rs("total_parts") = "0" then 
			strTotal = strTotal & "<h3>Total Parts: -</h3>"
		else
			strTotal = strTotal & "<h3>Total Parts: $" & FormatNumber(rs("total_parts")) & "</h3>"
		end if
		
		if IsNull(rs("total_gst")) or rs("total_gst") = "" or rs("total_gst") = "0" then
			strTotal = strTotal & "<h3>Total GST: -</h3>"
		else
			strTotal = strTotal & "<h3>Total GST: $" & FormatNumber(rs("total_gst")) & "</h3>"
		end if
		
		if IsNull(rs("total_total")) or rs("total_total") = "" or rs("total_total") = "0" then
			strTotal = strTotal & "<h2>Total Cost: -</h2>"
		else
			strTotal = strTotal & "<h2>Total Cost: $<u>" & FormatNumber(rs("total_total")) & "</u></h2>"
		end if
	else
        strTotal = "-"
	end if
	
    call CloseDataBase()
end function

%>