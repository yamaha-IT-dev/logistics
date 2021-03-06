<%@ Language=VBScript %>
<!--#include file="include/connection_it.asp " -->
<%
'Response.Write "hello"

dim rs
dim strSQL

dim strDays
dim strTodayDate
strTodayDate = FormatDateTime(Date())

Call OpenDataBase()

set rs=server.createobject("ADODB.recordset")

	strSQL = "SELECT * FROM yma_quarantines "
	strSQL = strSQL & "	WHERE department LIKE '%" & session("return_department") & "%' "	
	if session("return_type") <> "" then
		strSQL = strSQL & "		AND return_type = '" & session("return_type") & "' "
	end if	
	strSQL = strSQL & "		AND (item_code LIKE '%" & session("return_search") & "%' "
	strSQL = strSQL & "			OR shipment_no LIKE '%" & session("return_search") & "%' "
	strSQL = strSQL & "			OR description LIKE '%" & session("return_search") & "%' "
	strSQL = strSQL & "			OR return_carrier LIKE '%" & session("return_search") & "%' "
	strSQL = strSQL & "			OR return_connote LIKE '%" & session("return_search") & "%' "
	strSQL = strSQL & "			OR original_connote LIKE '%" & session("return_search") & "%' "
	strSQL = strSQL & "			OR gra LIKE '%" & session("return_search") & "%' "
	strSQL = strSQL & "			OR serial_no LIKE '%" & session("return_search") & "%' ) "
	strSQL = strSQL & "		AND instruction LIKE '%" & session("return_instruction") & "%' "
	strSQL = strSQL & "		AND status LIKE '%" & session("return_status") & "%' "
	strSQL = strSQL & "	ORDER BY date_created DESC"

rs.open strSQL,conn,1,3

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=warehouse-return_list.xls"

if rs.eof <> true then
	Response.Write "<table border=1>"
	Response.Write "<tr>"
	Response.Write "<td><strong>Type</strong></td>"
	Response.Write "<td><strong>Dept</strong></td>"
	Response.Write "<td><strong>Days</strong></td>"
	Response.Write "<td><strong>Item</strong></td>"
	Response.Write "<td><strong>Return connote</strong></td>"
	Response.Write "<td><strong>Dealer</strong></td>"
	Response.Write "<td><strong>Shipment no</strong></td>"
	Response.Write "<td><strong>Qty</strong></td>"		
	Response.Write "<td><strong>Carrier</strong></td>"
	Response.Write "<td><strong>Original connote</strong></td>"
	Response.Write "<td><strong>Serial no</strong></td>"
	Response.Write "<td><strong>Instruction</strong></td>"	
	Response.Write "<td><strong>GRA</strong></td>"
	Response.Write "<td><strong>Status</strong></td>"
	Response.Write "</tr>"   
	while not rs.eof
		strDays = DateDiff("d",rs("date_created"), strTodayDate)
		Response.Write "<tr>"
		
		Response.Write "<td>"
		Select Case rs.fields("return_type") 
			case 1 
				Response.Write "Managed"
			case 2
				Response.Write "Un-addressed"
			case 0
				Response.Write "Un-managed"
			case else
				Response.Write "-"
		end select
		Response.Write "</td>"   
		
		Response.Write "<td>" & rs.fields("department") & "</td>"
		Response.Write "<td>" & strDays & "</td>"
		Response.Write "<td>" & rs.fields("item_code") & "</td>"
		Response.Write "<td>" & rs.fields("return_connote") & "</td>"
		Response.Write "<td>" & rs.fields("dealer") & "</td>"
		Response.Write "<td>" & rs.fields("shipment_no") & "</td>"
		Response.Write "<td>" & rs.fields("qty") & "</td>"		
		Response.Write "<td>" & rs.fields("return_carrier") & "</td>"		
		Response.Write "<td>" & rs.fields("original_connote") & "</td>"
		Response.Write "<td>" & rs.fields("serial_no") & "</td>"
		Response.Write "<td>"
		Select Case rs.fields("instruction")
			case 1 
				Response.Write "Return to good stock 3T"
			case 2
				Response.Write "Transfer to Excel 3XL"
			case 3
				Response.Write "Resend to customer"
			case 4 
				Response.Write "Damaged item to Excel - good stock to 3T"		
			case else
				Response.Write "-"
		end select
		Response.Write "</td>"
		
		Response.Write "<td>" & rs.fields("gra") & "</td>"
		if rs.fields("status") = 1 then
			Response.Write "<td>Open</td>"
		else
			Response.Write "<td>Closed</td>"
		end if
		Response.Write "</tr>"
		rs.movenext
	wend
	Response.Write "</table>"
end if

Call CloseDataBase()
%>