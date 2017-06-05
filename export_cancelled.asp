<%@ Language=VBScript %>
<!--#include file="include/connection_it.asp " -->
<%
dim rs
dim strSQL

dim strSearch
strSearch 	= Trim(Request("search"))

dim intStatus
intStatus 	= Trim(Request("status"))

Call OpenDataBase()

set rs=server.createobject("ADODB.recordset")

strSQL = "SELECT * FROM yma_cancelled_order "
strSQL = strSQL & "	WHERE  (cancel_shipment_no LIKE '%" & strSearch & "%' "
strSQL = strSQL & "			OR cancel_info LIKE '%" & strSearch & "%' "
strSQL = strSQL & "			OR cancel_created_by LIKE '%" & strSearch & "%') "
strSQL = strSQL & "		AND cancel_status LIKE '%" & intStatus& "%' "
strSQL = strSQL & "	ORDER BY cancel_date_created DESC"

rs.open strSQL,conn,1,3

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=cancelled_order_list.xls"

if rs.eof <> true then
	Response.Write "<table border=1>"
	Response.Write "<tr>"
	Response.Write "<td><strong>Created by</strong></td>"
	Response.Write "<td><strong>Date Created</strong></td>"
	Response.Write "<td><strong>Shipment no</strong></td>"
	Response.Write "<td><strong>Info</strong></td>"
	Response.Write "<td><strong>Warehouse Confirm</strong></td>"	
	Response.Write "<td><strong>Logistics Confirm</strong></td>"
	Response.Write "<td><strong>Status</strong></td>"
	Response.Write "<td><strong>Last Modified by</strong></td>"
	Response.Write "<td><strong>Last Modified Date</strong></td>"	
	Response.Write "</tr>"   
	while not rs.eof
		Response.Write "<tr>"
		Response.Write "<td>" & rs.fields("cancel_created_by") & "</td>"
		Response.Write "<td>" & rs.fields("cancel_date_created") & "</td>"
		Response.Write "<td>" & rs.fields("cancel_shipment_no") & "</td>"
		Response.Write "<td>" & rs.fields("cancel_info") & "</td>"
		if rs.fields("cancel_warehouse_confirm") = 1 then
			Response.Write "<td>Y</td>"
		else
			Response.Write "<td>N</td>"
		end if
		if rs.fields("cancel_logistics_confirm") = 1 then
			Response.Write "<td>Y</td>"
		else
			Response.Write "<td>N</td>"
		end if
		
		if rs.fields("cancel_status") = 1 then
			Response.Write "<td>open</td>"
		else
			Response.Write "<td>completed</td>"
		end if
		
		Response.Write "<td>" & rs.fields("cancel_modified_by") & "</td>"
		Response.Write "<td>" & rs.fields("cancel_date_modified") & "</td>"				
		Response.Write "</tr>"
		rs.movenext
	wend
	Response.Write "</table>"
end if

Call CloseDataBase()
%>