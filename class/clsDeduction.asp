<%
'----------------------------------------------------------------------------------------
' GET ALL DEDUCTION TYPE LIST: view_gra.asp
'----------------------------------------------------------------------------------------
function getAllDeductionTypeList(intGraNo)
    dim strSQL
    dim rs
	dim intDeductionTypeID
	dim strDeductionTypeName
	dim intDeductionTypeAmount
	
    call OpenDataBase()
    
	strSQL = "SELECT DT.* FROM tbl_deduction_type DT "
	strSQL = strSQL & "	WHERE deduction_type_status = 1 "
	'strSQL = strSQL & "		AND DT.deduction_type_id NOT IN (SELECT D.deduction_type_id FROM tbl_deductions D WHERE deduction_gra_no = '" & intGraNo & "')"
	strSQL = strSQL & "	ORDER BY deduction_type_name"
		
	set rs = server.CreateObject("ADODB.Recordset")
    set rs = conn.execute(strSQL)
    
    strAllDeductionTypeList = strAllDeductionTypeList & "<option value=''>...</option>"
    
    if not DB_RecSetIsEmpty(rs) Then
        do until rs.EOF
        	intDeductionTypeID		= trim(rs("deduction_type_id"))
			strDeductionTypeName	= trim(rs("deduction_type_name"))
			intDeductionTypeAmount	= trim(rs("deduction_type_amount"))
			
			strAllDeductionTypeList = strAllDeductionTypeList & "<option value=" & intDeductionTypeID & ">" & strDeductionTypeName & " - $" & FormatNumber(intDeductionTypeAmount) & "</option>"
        rs.Movenext
        loop
    end if
    
    call CloseDataBase()
end function

'-----------------------------------------------
' LIST DEDUCTIONS
'-----------------------------------------------
function listDeductions(intGraNo)
    dim strSQL
	dim intRecordCount
	dim intTotalDeduction
	intTotalDeduction = 0
	
    call OpenDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic			
	rs.PageSize = 200
	
	strSQL = "SELECT D.*, DT.deduction_type_name, DT.deduction_type_amount"
	'strSQL = strSQL & ", SUM(deduction_type_amount) as deduction_total "
	strSQL = strSQL & "		FROM tbl_deductions D "
	strSQL = strSQL & " 		INNER JOIN tbl_deduction_type DT on D.deduction_type_id = DT.deduction_type_id "
	strSQL = strSQL & "				WHERE D.deduction_gra_no = '" & intGraNo & "'"
	strSQL = strSQL & "			ORDER BY D.deduction_line, DT.deduction_type_name"
	'strSQL = strSQL & "			GROUP BY D.deduction_gra_no"
	
	rs.Open strSQL, conn
	
	'response.write strSQL
	
	intRecordCount = rs.recordcount

    strDeductionList = ""
	
	if not DB_RecSetIsEmpty(rs) Then	
	
		For intRecord = 1 To rs.PageSize
		
			strDeductionList = strDeductionList & "<tr>"
			strDeductionList = strDeductionList & "<td>" & trim(rs("deduction_line")) & "</td>"
			strDeductionList = strDeductionList & "<td>" & trim(rs("deduction_type_name")) & "</td>"
			strDeductionList = strDeductionList & "<td>" & FormatNumber(rs("deduction_type_amount")) & "</td>"
			strDeductionList = strDeductionList & "<td>" & rs("deduction_comments") & "</td>"
			strDeductionList = strDeductionList & "<td>" & trim(rs("deduction_created_by")) & "</td>"
			strDeductionList = strDeductionList & "<td>" & trim(rs("deduction_date_created")) & "</td>"
			strDeductionList = strDeductionList & "<td><a onclick=""return confirm('Are you sure you want to delete " & rs("deduction_type_name") & "?');"" href='delete_deduction.asp?gra_no=" & trim(rs("deduction_gra_no")) & "&id=" & trim(rs("deduction_id")) & "'><img src=""images/btn_delete.gif"" border=""0""></a></td>"
			strDeductionList = strDeductionList & "</tr>"
			'intTotalDeduction = intTotalDeduction + trim(rs("deduction_amount"))
			
			rs.movenext
			
			If rs.EOF Then Exit For
		next
	else
        strDeductionList = "<tr><td>&nbsp;</td></tr>"
	end if
	'strDeductionList = strDeductionList & "<tr>"
	'strDeductionList = strDeductionList & "		<td width=""30%""><strong>Total Deduction:</strong></td>"
	'strDeductionList = strDeductionList & "		<td width=""70%"" colspan=""4""><strong>$" & FormatNumber(deduction_total) & "</strong></td>"
	'strDeductionList = strDeductionList & "</tr>"
	'strDeductionList = strDeductionList & "<tr>"
	
    call CloseDataBase()
end function

'-----------------------------------------------
' SUM DEDUCTIONS
'-----------------------------------------------
function sumTotalDeductions(intGraNo)
    dim strSQL
	dim intTotalDeduction
	
    call OpenDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic			
	rs.PageSize = 200
	
	strSQL = "SELECT SUM(deduction_type_amount) as deduction_total "
	strSQL = strSQL & "		FROM tbl_deductions D "
	strSQL = strSQL & " 		INNER JOIN tbl_deduction_type DT on D.deduction_type_id = DT.deduction_type_id "
	strSQL = strSQL & "				WHERE D.deduction_gra_no = '" & intGraNo & "'"
	
	rs.Open strSQL, conn
	
	'response.write strSQL
	intTotalDeduction = 0
	
	if not DB_RecSetIsEmpty(rs) Then	
		if not isnull(rs("deduction_total")) then		
			intTotalDeduction = rs("deduction_total")
			response.write (FormatNumber(intTotalDeduction))
		else
			response.Write ("0")
		end if
	end if
	
    call CloseDataBase()
end function

'-----------------------------------------------
' ADD DEDUCTION
'-----------------------------------------------
function addDeduction(intGraNo,intDeductionTypeID,intDeductionLine,strDeductionComments)
	dim strSQL
	
	Call OpenDataBase()
		
	strSQL = "INSERT INTO tbl_deductions ("
	strSQL = strSQL & "deduction_gra_no,"
	strSQL = strSQL & "deduction_type_id,"
	strSQL = strSQL & "deduction_line,"
	strSQL = strSQL & "deduction_comments,"
	strSQL = strSQL & "deduction_created_by"
	strSQL = strSQL & ") VALUES ("
	strSQL = strSQL & "'" & trim(intGraNo) & "',"
	strSQL = strSQL & "'" & trim(intDeductionTypeID) & "',"
	strSQL = strSQL & "'" & trim(intDeductionLine) & "',"
	strSQL = strSQL & "'" & Server.HTMLEncode(strDeductionComments) & "',"
	strSQL = strSQL & "'" & session("UsrUserName") & "')"

	on error resume next
	conn.Execute strSQL
	
	'response.write strSQL
	
	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "Deduction has been added."
	end if
	
	Call CloseDataBase()
end function
%>