<%@ Language=VBScript %>
<!--#include file="include/connection_it.asp " -->
<%
dim rs
dim sql

dim strSearch
strSearch 	= Trim(Request("search"))

dim strState
strState = Trim(Request("state"))

dim intStatus
intStatus 	= Trim(Request("status"))

Call OpenDataBase()

set rs=server.createobject("ADODB.recordset")

sql = "SELECT * FROM yma_changeover WHERE state LIKE '%" & strState & "%' AND (customer LIKE '%" & strSearch & "%' OR contact_person LIKE '%" & strSearch & "%' OR old_model LIKE '%" & strSearch & "%') AND status LIKE '%" & intStatus & "%'"

rs.open sql,conn,1,3

'on error resume next
'conn.Execute sql

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=changeover_list.xls"

if rs.eof <> true then
	Response.Write "<table border=1>"
	Response.Write "<tr>"
	Response.Write "<td><strong>Claim no</strong></td>"
	Response.Write "<td><strong>Customer</strong></td>"
	Response.Write "<td><strong>Contact person</strong></td>"
	Response.Write "<td><strong>Primary contact no</strong></td>"	
	Response.Write "<td><strong>Secondary contact no</strong></td>"
	Response.Write "<td><strong>Address</strong></td>"
	Response.Write "<td><strong>City</strong></td>"
	Response.Write "<td><strong>State</strong></td>"
	Response.Write "<td><strong>Postcode</strong></td>"	
	Response.Write "<td><strong>Old model</strong></td>"
	Response.Write "<td><strong>Old model serial</strong></td>"
	Response.Write "<td><strong>Proof</strong></td>"
	Response.Write "<td><strong>Warranty</strong></td>"
	Response.Write "<td><strong>Replacement</strong></td>"
	Response.Write "<td><strong>Make up cost</strong></td>"
	Response.Write "<td><strong>Replacement going to</strong></td>"
	Response.Write "<td><strong>Date Received</strong></td>"
	Response.Write "<td><strong>Date Paid</strong></td>"
	Response.Write "<td><strong>Invoice no</strong></td>"
	Response.Write "<td><strong>Status</strong></td>"
	Response.Write "</tr>"   
	while not rs.eof
		Response.Write "<tr>"
		Response.Write "<td>" & rs.fields("changeover_id") & "</td>"
		Response.Write "<td>" & rs.fields("customer") & "</td>"
		Response.Write "<td>" & rs.fields("contact_person") & "</td>"
		Response.Write "<td>" & rs.fields("phone") & "</td>"
		Response.Write "<td>" & rs.fields("mobile") & "</td>"
		Response.Write "<td>" & rs.fields("address") & "</td>"
		Response.Write "<td>" & rs.fields("city") & "</td>"
		Response.Write "<td>" & rs.fields("state") & "</td>"
		Response.Write "<td>" & rs.fields("postcode") & "</td>"
		Response.Write "<td>" & rs.fields("old_model") & "</td>"
		Response.Write "<td>" & rs.fields("old_model_serial") & "</td>"		
		'Response.Write "<td>" & rs.fields("proof") & "</td>"
		if rs.fields("proof") = 1 then
			Response.Write "<td>Y</td>"
		else
			Response.Write "<td>N</td>"
		end if
		'Response.Write "<td>" & rs.fields("warranty") & "</td>"
		if rs.fields("warranty") = 1 then
			Response.Write "<td>Y</td>"
		else
			Response.Write "<td>N</td>"
		end if
		Response.Write "<td>" & rs.fields("replacement_model") & "</td>"
		Response.Write "<td>$" & rs.fields("make_up_cost") & "</td>"
		Response.Write "<td>" & rs.fields("replacement_destination") & "</td>"
		Response.Write "<td>" & rs.fields("date_received") & "</td>"
		Response.Write "<td>" & rs.fields("date_payment") & "</td>"
		Response.Write "<td>" & rs.fields("invoice_no") & "</td>"
		
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