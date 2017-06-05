<%@ Language=VBScript %>
<!--#include file="include/connection_it.asp " -->
<%
dim rs
dim strSQL

dim strSearch
strSearch 	= Trim(Request("search"))

dim strPallet
strPallet 	= Trim(Request("pallet"))

dim intYear
intYear 	= Trim(Request("year"))

dim intMonth
intMonth 	= Trim(Request("month"))

dim intStatus
intStatus 	= Trim(Request("status"))

dim strSort
strSort 	= Trim(Request("sort"))

if strSort = "" then
	strSort = "gra"
end if

if intYear = "" then
	intYear = 2015
end if

Call OpenDataBase()

set rs = server.createobject("ADODB.recordset")

	strSQL = "SELECT created_by, date_created, gra_no, line_no, item, serial_no, LIC, dealer_code, repair_report, labour, parts, gst, total, date_received, date_repaired, destination, pallet_no, status, LIC FROM yma_gra_report "	
	strSQL = strSQL & " LEFT JOIN "
	strSQL = strSQL & "		OPENQUERY(AS400, 'SELECT E2SOSC, E2NGTY, E2NGTM, (E2ihtn+(e2ihtn*e2kzrt/100)+(e2ihtn*e2skkr/100)) as LIC FROM EF2SP') "
	strSQL = strSQL & "				ON item = E2SOSC and E2NGTY = YEAR(getdate()) AND E2NGTM = MONTH(getdate()) "
	strSQL = strSQL & " WHERE pallet_no LIKE '%" & strPallet & "%' "
	
	if intYear <> "" then
		strSQL = strSQL & " AND YEAR(" & session("gra_writeoff_report_date") & ") = '" & intYear & "' "
	end if
	
	if intMonth <> "" then
		strSQL = strSQL & " AND MONTH(" & session("gra_writeoff_report_date") & ") = '" & intMonth & "' "
	end if	
	
	strSQL = strSQL & " 	AND (gra_no LIKE '%" & strSearch & "%' "
	strSQL = strSQL & "			OR item LIKE '%" & strSearch & "%' "
	strSQL = strSQL & "			OR dealer_code LIKE '%" & strSearch & "%' "
	strSQL = strSQL & "			OR pallet_no LIKE '%" & strSearch & "%' "
	strSQL = strSQL & "			OR serial_no LIKE '%" & strSearch & "%') "
	strSQL = strSQL & " 	AND status LIKE '%" & intStatus & "%' "
	strSQL = strSQL & " 	AND destination = 'Destroy'"	
	strSQL = strSQL & " ORDER BY " 
	select case strSort
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
Response.AddHeader "Content-Disposition", "attachment; filename=gra-writeoff-report-"& intYear & "_list.xls"

if rs.eof <> true then
	Response.Write "<table border=1>"
	Response.Write "<tr>"
	Response.Write "<td nowrap><strong>Created by</strong></td>"
	Response.Write "<td nowrap><strong>Date created</strong></td>"
	Response.Write "<td nowrap><strong>GRA no</strong></td>"
	Response.Write "<td nowrap><strong>Line</strong></td>"
	Response.Write "<td nowrap><strong>Item</strong></td>"	
	Response.Write "<td nowrap><strong>Serial no</strong></td>"
	Response.Write "<td nowrap><strong>LIC</strong></td>"
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
		Response.Write "<td nowrap>" & rs.fields("line_no") & "</td>"
		Response.Write "<td nowrap>" & rs.fields("item") & "</td>"
		Response.Write "<td nowrap>" & rs.fields("serial_no") & "</td>"
		Response.Write "<td nowrap>" & FormatNumber(rs.fields("LIC")) & "</td>"
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
				Response.Write "<td>Completed / Exported</td>"
		end select
		Response.Write "</tr>"
		
		rs.movenext
	wend	
	Response.Write "</table>"
end if

Call CloseDataBase()
%>