<%@ Language=VBScript %>
<!--#include file="include/connection_it.asp " -->
<%
dim rs
dim strSQL

if session("gra_report_date") = "" then
	session("gra_report_date") = "date_created"
end if

if session("gra_report_sort") = "" then
	session("gra_report_sort") = "gra"
end if

Call OpenDataBase()

set rs = server.createobject("ADODB.recordset")

	strSQL = "SELECT yma_gra_report.*, BTAHYD, BTAHYM, BTAHYY FROM yma_gra_report "
	strSQL = strSQL & " LEFT JOIN OPENQUERY(AS400, 'SELECT BTHYNO, BTAHYD, BTAHYM, BTAHYY FROM BFTEP') ON BTHYNO = gra_no "
	strSQL = strSQL & " WHERE "
	strSQL = strSQL & " 	(gra_no LIKE '%" & session("gra_report_search") & "%' "
	strSQL = strSQL & "			OR dealer_code LIKE '%" & session("gra_report_search") & "%' "
	strSQL = strSQL & "			OR item LIKE '%" & session("gra_report_search") & "%' "
	strSQL = strSQL & "			OR serial_no LIKE '%" & session("gra_report_search") & "%') "
	
	if session("gra_report_month") <> "" then
		strSQL = strSQL & " AND MONTH(" & session("gra_report_date") & ") = '" & trim(session("gra_report_month")) & "' "
	end if
	
	if session("gra_report_year") <> "" then
		strSQL = strSQL & " AND YEAR(" & session("gra_report_date") & ") = '" & trim(session("gra_report_year")) & "' "
	end if
		
	strSQL = strSQL & " 	AND status LIKE '%" & session("gra_report_status") & "%' "
	strSQL = strSQL & " ORDER BY " 
	select case session("gra_report_sort")
		case "gra"
			strSQL = strSQL & "gra_no"
		case "item"
			strSQL = strSQL & "item"
		case "created_latest"
			strSQL = strSQL & "date_created DESC"
		case "created_oldest"
			strSQL = strSQL & "date_created"	
		case "received_latest"
			strSQL = strSQL & "date_received DESC"
		case "received_oldest"
			strSQL = strSQL & "date_received"
		case "repaired_latest"
			strSQL = strSQL & "date_repaired DESC"
		case "repaired_oldest"
			strSQL = strSQL & "date_repaired"	
		case "exported_latest"
			strSQL = strSQL & "date_exported DESC"
		case "exported_oldest"
			strSQL = strSQL & "date_exported"	
		case else
			strSQL = strSQL & "gra_no"
	end select

rs.open strSQL,conn,1,3

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=gra-report.xls"

if rs.eof <> true then
	Response.Write "<table border=1>"
	Response.Write "<tr>"
	Response.Write "<td nowrap><strong>Created by</strong></td>"
	Response.Write "<td nowrap><strong>Date created</strong></td>"
	Response.Write "<td nowrap><strong>GRA no</strong></td>"
	Response.Write "<td nowrap><strong>GRA generated</strong></td>"
	Response.Write "<td nowrap><strong>Line</strong></td>"
	Response.Write "<td nowrap><strong>Item</strong></td>"	
	Response.Write "<td nowrap><strong>Warranty Code</strong></td>"	
	Response.Write "<td nowrap><strong>Serial no</strong></td>"
	Response.Write "<td nowrap><strong>Dealer code</strong></td>"
	Response.Write "<td nowrap><strong>Repair report</strong></td>"
	Response.Write "<td nowrap><strong>Labour</strong></td>"
	Response.Write "<td nowrap><strong>Parts</strong></td>"
	Response.Write "<td nowrap><strong>GST</strong></td>"
	Response.Write "<td nowrap><strong>Total</strong></td>"
	Response.Write "<td nowrap><strong>Received</strong></td>"
	Response.Write "<td nowrap><strong>Repaired</strong></td>"
	Response.Write "<td nowrap><strong>Destination</strong></td>"
	Response.Write "<td nowrap><strong>Pallet no</strong></td>"
	Response.Write "<td nowrap><strong>Report status</strong></td>"
	Response.Write "</tr>"   
	while not rs.eof
		Response.Write "<tr>"
		Response.Write "<td nowrap>" & rs.fields("created_by") & "</td>"
		Response.Write "<td nowrap>" & rs.fields("date_created") & "</td>"
		Response.Write "<td nowrap>" & rs.fields("gra_no") & "</td>"
		Response.Write "<td nowrap>" & rs.fields("BTAHYD") & "/" & rs.fields("BTAHYM") & "/" & rs.fields("BTAHYY") & "</td>"
		Response.Write "<td nowrap>" & rs.fields("line_no") & "</td>"
		Response.Write "<td nowrap>" & rs.fields("item") & "</td>"
		Response.Write "<td nowrap>" & rs.fields("gra_warranty_code") & "</td>"
		Response.Write "<td nowrap>" & rs.fields("serial_no") & "</td>"
		Response.Write "<td nowrap>" & rs.fields("dealer_code") & "</td>"
		Response.Write "<td nowrap>" & rs.fields("repair_report") & "</td>"
		Response.Write "<td nowrap>" & rs.fields("labour") & "</td>"
		Response.Write "<td nowrap>" & rs.fields("parts") & "</td>"
		Response.Write "<td nowrap>" & rs.fields("gst") & "</td>"
		Response.Write "<td nowrap>" & rs.fields("total") & "</td>"
		Response.Write "<td nowrap>" & rs.fields("date_received") & "</td>"
		Response.Write "<td nowrap>" & rs.fields("date_repaired") & "</td>"
		Response.Write "<td nowrap>" & rs.fields("destination") & "</td>"
		Response.Write "<td nowrap>" & rs.fields("pallet_no") & "</td>"
		select case rs.fields("status") 
			case 1 
				Response.Write "<td>Open</td>"
			case 2
				Response.Write "<td>Waiting for parts</td>"
			case 3
				Response.Write "<td>To be invoiced</td>"
			case 4
				Response.Write "<td>Received</td>"
			case else
				Response.Write "<td>Completed / Exported - " & rs.fields("date_exported") & "</td>"
		end select
		Response.Write "</tr>"
				
		rs.movenext
	wend	
	Response.Write "</table>"
end if

Call CloseDataBase()
%>