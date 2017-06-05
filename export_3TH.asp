<%@ Language=VBScript %>
<!--#include file="include/connection_it.asp " -->
<%
'Response.Write "hello"

dim rs
dim strSQL

dim strDays
dim strTodayDate
strTodayDate = FormatDateTime(Date())

dim strDepartment
strDepartment 	= Trim(Request("dept"))

dim strSearch
strSearch 	= Trim(Request("search"))

dim intType
intType 	= Trim(Request("type"))

dim intStatus
intStatus 	= Trim(Request("status"))

Call OpenDataBase()

set rs=server.createobject("ADODB.recordset")

strSQL = "SELECT * FROM yma_3th_return "
strSQL = strSQL & "	WHERE department LIKE '%" & strDepartment & "%' "
	
if intType <> "" then
	strSQL = strSQL & "		AND return_type = '" & intType & "' "
end if
	
strSQL = strSQL & "		AND (item_code LIKE '%" & strSearch & "%' "
strSQL = strSQL & "			OR shipment_no LIKE '%" & strSearch & "%' "
strSQL = strSQL & "			OR description LIKE '%" & strSearch & "%' "
strSQL = strSQL & "			OR carrier LIKE '%" & strSearch & "%' "
strSQL = strSQL & "			OR label_no LIKE '%" & strSearch & "%' "
strSQL = strSQL & "			OR original_connote LIKE '%" & strSearch & "%' "
strSQL = strSQL & "			OR gra LIKE '%" & strSearch & "%' "
strSQL = strSQL & "			OR serial_no LIKE '%" & strSearch & "%' ) "
strSQL = strSQL & "		AND status LIKE '%" & intStatus & "%' "
strSQL = strSQL & "	ORDER BY date_created DESC"

rs.open strSQL,conn,1,3

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=3TH_list.xls"

if rs.eof <> true then
	Response.Write "<table border=1>"
	Response.Write "<tr>"
		  Response.Write "<td><strong>Type</strong></td>"
          Response.Write "<td><strong>Dept</strong></td>"
          Response.Write "<td><strong>Lapsed</strong></td>"
          Response.Write "<td><strong>Item</strong></td>"
          Response.Write "<td><strong>Qty</strong></td>"
          Response.Write "<td><strong>Description</strong></td>"
          Response.Write "<td><strong>Label #</strong></td>"
          Response.Write "<td><strong>Dealer</strong></td>"
          Response.Write "<td><strong>Shipment #</strong></td>"         
          Response.Write "<td><strong>Carrier</strong></td>"
          Response.Write "<td><strong>Connote</strong></td>"
          Response.Write "<td><strong>Received</strong></td>"
          Response.Write "<td><strong>Serial #</strong></td>"
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
					Response.Write "Lost in Warehouse"
				case 2
					Response.Write "Lost by Carrier"
				case 3
					Response.Write "Packaging Issue"
				case 4
					Response.Write "Warehouse Variance"
				case 5
					Response.Write "Display Stock"
				case else
					Response.Write "-"
			end select
			Response.Write "</td>"
			Response.Write "<td>" & rs.fields("department") & "</td>"
			Response.Write "<td>" & strDays & "</td>"
			Response.Write "<td>" & rs.fields("item_code") & "</td>"
			Response.Write "<td>" & rs.fields("qty") & "</td>"
			Response.Write "<td>" & rs.fields("description") & "</td>"
			Response.Write "<td>" & rs.fields("label_no") & "</td>"
			Response.Write "<td>" & rs.fields("dealer") & "</td>"
			Response.Write "<td>" & rs.fields("shipment_no") & "</td>"					
			Response.Write "<td>" & rs.fields("carrier") & "</td>"		
			Response.Write "<td>" & rs.fields("original_connote") & "</td>"
			Response.Write "<td>" & rs.fields("date_received") & "</td>"
			Response.Write "<td>" & rs.fields("serial_no") & "</td>"
			Response.Write "<td>"
			Select Case rs.fields("instruction")
				case 1 
					Response.Write "Update GRA"
				case 2
					Response.Write "Writeoff Approval Required"				
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