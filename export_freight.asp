<%@ Language=VBScript %>
<!--#include file="include/connection_it.asp " -->
<%
dim rs
dim sql
dim strState
strState = Trim(Request("state"))

dim strSearch
strSearch 	= Trim(Request("search"))

dim intStatus
intStatus 	= Trim(Request("status"))

Call OpenDataBase()

set rs=server.createobject("ADODB.recordset")

sql =  "SELECT * FROM yma_freight WHERE receiver_state LIKE '%" & strState & "%' AND (receiver_name LIKE '%" & strSearch & "%' OR receiver_address LIKE '%" & session("freight_search") & "%' OR items LIKE '%" & strSearch & "%') AND status LIKE '%" & intStatus & "%' ORDER BY date_created DESC"

rs.open sql,conn,1,3

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=freight_list.xls"

if rs.eof <> true then
	Response.Write "<table border=1>"
	Response.Write "<tr>"
	Response.Write "<td><strong>ID</strong></td>"
	Response.Write "<td><strong>Requested By</strong></td>"
	Response.Write "<td><strong>Pickup</strong></td>"
	Response.Write "<td><strong>Phone</strong></td>"
	Response.Write "<td><strong>Address</strong></td>"	
	Response.Write "<td><strong>Date / Time</strong></td>"
	Response.Write "<td><strong>Return to pickup</strong></td>"
	Response.Write "<td><strong>Receiver</strong></td>"
	Response.Write "<td><strong>Phone</strong></td>"
	Response.Write "<td><strong>Address</strong></td>"
	Response.Write "<td><strong>Date / Time</strong></td>"
	Response.Write "<td><strong>Pickup</strong></td>"
	Response.Write "<td><strong>Connote</strong></td>"
	Response.Write "<td><strong>Return Connote</strong></td>"
	Response.Write "<td><strong>Status</strong></td>"
	Response.Write "</tr>"   
	while not rs.eof
		Response.Write "<tr>"
		Response.Write "<td>" & rs.fields("id") & "</td>"
		Response.Write "<td>" & rs.fields("email") & "</td>"
		Response.Write "<td>" & rs.fields("pickup_name") & "</td>"
		Response.Write "<td>" & rs.fields("pickup_phone") & "</td>"		
		Response.Write "<td>" & rs.fields("pickup_address") & " - " & rs.fields("pickup_city") & " " & rs.fields("pickup_state") & " " & rs.fields("pickup_postcode") & "</td>"
		Response.Write "<td>" & rs.fields("pickup_date") & "</td>"
		if rs.fields("return_pickup") = 1 then
			Response.Write "<td>Y</td>"
		else
			Response.Write "<td>N</td>"
		end if
		Response.Write "<td>" & rs.fields("receiver_name") & "</td>"
		Response.Write "<td>" & rs.fields("receiver_phone") & "</td>"
		Response.Write "<td>" & rs.fields("receiver_address") & " - " & rs.fields("receiver_city") & " " & rs.fields("receiver_state") & " " & rs.fields("receiver_postcode") & "</td>"
		Response.Write "<td>" & rs.fields("delivery_date") & " - " & rs.fields("delivery_time") & "</td>"
		if rs.fields("pickup") = 1 then
			Response.Write "<td>Y</td>"
		else
			Response.Write "<td>N</td>"
		end if
		Response.Write "<td>" & rs.fields("consignment_no") & "</td>"
		Response.Write "<td>" & rs.fields("return_connote") & "</td>"	
		if rs.fields("status") = 1 then
			Response.Write "<td>open</td>"
		else
			Response.Write "<td>completed</td>"
		end if
		Response.Write "</tr>"
		rs.movenext
	wend
	Response.Write "</table>"
end if

Call CloseDataBase()
%>