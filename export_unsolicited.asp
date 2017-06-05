<%@ Language=VBScript %>
<!--#include file="include/connection_it.asp " -->
<%
dim rs
dim strSQL

Call OpenDataBase()

set rs=server.createobject("ADODB.recordset")

	strSQL = "SELECT * FROM logistic_unsolicited "
	strSQL = strSQL & "	WHERE unsDepartment LIKE '%" & session("unsolicited_department") & "%' "	
	if session("unsolicited_instruction") <> "" then
		strSQL = strSQL & "		AND unsInstruction = '" & session("unsolicited_instruction") & "' "
	end if
	strSQL = strSQL & "		AND (unsItemCode LIKE '%" & session("unsolicited_search") & "%' "	
	strSQL = strSQL & "			OR unsDescription LIKE '%" & session("unsolicited_search") & "%' "
	strSQL = strSQL & "			OR unsConnote LIKE '%" & session("unsolicited_search") & "%' "
	strSQL = strSQL & "			OR unsGRA LIKE '%" & session("unsolicited_search") & "%' "
	strSQL = strSQL & "			OR unsDealer LIKE '%" & session("unsolicited_search") & "%' "	
	strSQL = strSQL & "			OR unsShipmentNo LIKE '%" & session("unsolicited_search") & "%' "
	strSQL = strSQL & "			OR unsComments LIKE '%" & session("unsolicited_search") & "%')"
	strSQL = strSQL & "		AND unsStatus LIKE '%" & session("unsolicited_status") & "%' "
	strSQL = strSQL & "	ORDER BY unsDateCreated DESC"

rs.open strSQL,conn,1,3

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=unsolicited_list.xls"

if rs.eof <> true then
	Response.Write "<table border=1>"
	Response.Write "<tr>"
	Response.Write "<td><strong>Created</strong></td>"
	Response.Write "<td><strong>Dept</strong></td>"
	Response.Write "<td><strong>Item</strong></td>"
	Response.Write "<td><strong>Description</strong></td>"
	Response.Write "<td><strong>Connote</strong></td>"
	Response.Write "<td><strong>GRA</strong></td>"
	Response.Write "<td><strong>Dealer</strong></td>"		
	Response.Write "<td><strong>Shipment</strong></td>"
	Response.Write "<td><strong>Qty</strong></td>"
	Response.Write "<td><strong>Instruction</strong></td>"	
	Response.Write "<td><strong>Comments</strong></td>"
	Response.Write "<td><strong>Status</strong></td>"
	Response.Write "</tr>"   
	while not rs.eof		
		Response.Write "<tr>"	
		Response.Write "<td>" & rs.fields("unsCreatedBy") & " - " & rs.fields("unsDateCreated") & "</td>"
		Response.Write "<td>" & rs.fields("unsDepartment") & "</td>"
		Response.Write "<td>" & rs.fields("unsItemCode") & "</td>"
		Response.Write "<td>" & rs.fields("unsDescription") & "</td>"
		Response.Write "<td>" & rs.fields("unsConnote") & "</td>"
		Response.Write "<td>" & rs.fields("unsGRA") & "</td>"
		Response.Write "<td>" & rs.fields("unsDealer") & "</td>"		
		Response.Write "<td>" & rs.fields("unsShipmentNo") & "</td>"		
		Response.Write "<td>" & rs.fields("unsQty") & "</td>"
		Response.Write "<td>"
		Select Case rs.fields("unsInstruction")
			case 1 
				Response.Write "Move to 3XL"
			case 2
				Response.Write "Move to 3S"
			case 3
				Response.Write "Investigate"			
		end select
		Response.Write "</td>"		
		Response.Write "<td>" & rs.fields("unsComments") & "</td>"
		if rs.fields("unsStatus") = 1 then
			Response.Write "<td>Open</td>"
		else
			Response.Write "<td>Completed</td>"
		end if
		Response.Write "</tr>"
		rs.movenext
	wend
	Response.Write "</table>"
end if

Call CloseDataBase()
%>