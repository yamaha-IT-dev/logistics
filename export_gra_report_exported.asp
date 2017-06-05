<%@ Language=VBScript %>
<!--#include file="include/connection_it.asp " -->
<%
dim rs
dim strSQL

dim strSearch
strSearch 	= Trim(Request("search"))

dim strPallet
strPallet 	= Trim(Request("pallet"))

dim strStart
strStart 	= Trim(Request("start"))

dim strEnd
strEnd 		= Trim(Request("end"))

dim strSort
strSort 	= Trim(Request("sort"))

if strSort = "" then
	strSort = "gra_no"
end if

Call OpenDataBase()

set rs = server.createobject("ADODB.recordset")

strSQL = "SELECT yma_gra_report.*, BTAHYD, BTAHYM, BTAHYY FROM yma_gra_report "
strSQL = strSQL & " LEFT JOIN OPENQUERY(AS400, 'SELECT BTHYNO, BTAHYD, BTAHYM, BTAHYY FROM BFTEP') ON BTHYNO = gra_no "
strSQL = strSQL & " WHERE (date_exported BETWEEN CONVERT(datetime,'" & strStart & " 00:00:00',103)" & " "
strSQL = strSQL & " 	AND CONVERT(datetime,'" & strEnd & " 23:59:59',103)" & ") "
strSQL = strSQL & " 	AND (gra_no LIKE '%" & strSearch & "%' "
strSQL = strSQL & "			OR item LIKE '%" & strSearch & "%') "
strSQL = strSQL & " 	AND status = '0' "
strSQL = strSQL & " ORDER BY " & strSort

rs.open strSQL,conn,1,3

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=gra-exported-report-"& strStart & "_list.xls"

if rs.eof <> true then
	Response.Write "<table border=1>"
	Response.Write "<tr>"
	Response.Write "<td><strong>Created by</strong></td>"
	Response.Write "<td><strong>Date created</strong></td>"
	Response.Write "<td><strong>GRA no</strong></td>"
	Response.Write "<td nowrap><strong>GRA generated</strong></td>"
	Response.Write "<td><strong>Line</strong></td>"
	Response.Write "<td><strong>Item</strong></td>"	
	Response.Write "<td><strong>Serial no</strong></td>"
	Response.Write "<td><strong>Dealer code</strong></td>"
	Response.Write "<td><strong>Repair report</strong></td>"
	Response.Write "<td><strong>Labour</strong></td>"
	Response.Write "<td><strong>Parts</strong></td>"
	Response.Write "<td><strong>GST</strong></td>"
	Response.Write "<td><strong>Total</strong></td>"
	Response.Write "<td><strong>Received</strong></td>"
	Response.Write "<td><strong>Repaired</strong></td>"
	Response.Write "<td><strong>Destination</strong></td>"
	Response.Write "<td><strong>Date exported</strong></td>"
	Response.Write "<td><strong>Report status</strong></td>"
	Response.Write "</tr>"   
	while not rs.eof
		Response.Write "<tr>"
		Response.Write "<td>" & rs.fields("created_by") & "</td>"
		Response.Write "<td>" & rs.fields("date_created") & "</td>"
		Response.Write "<td>" & rs.fields("gra_no") & "</td>"
		Response.Write "<td nowrap>" & rs.fields("BTAHYD") & "/" & rs.fields("BTAHYM") & "/" & rs.fields("BTAHYY") & "</td>"
		Response.Write "<td>" & rs.fields("line_no") & "</td>"
		Response.Write "<td>" & rs.fields("item") & "</td>"
		Response.Write "<td>" & rs.fields("serial_no") & "</td>"
		Response.Write "<td>" & rs.fields("dealer_code") & "</td>"
		Response.Write "<td>" & rs.fields("repair_report") & "</td>"
		Response.Write "<td>" & rs.fields("labour") & "</td>"
		Response.Write "<td>" & rs.fields("parts") & "</td>"
		Response.Write "<td>" & rs.fields("gst") & "</td>"
		Response.Write "<td>" & rs.fields("total") & "</td>"	
		Response.Write "<td>" & rs.fields("date_received") & "</td>"
		Response.Write "<td>" & rs.fields("date_repaired") & "</td>"
		Response.Write "<td>" & rs.fields("destination") & "</td>"
		Response.Write "<td>" & rs.fields("date_exported") & "</td>"
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