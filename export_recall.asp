<%@ Language=VBScript %>
<!--#include file="include/connection_it.asp " -->
<%
dim rs
dim sql
dim strDept
strDept = Trim(Request("cboDepartment"))

Call OpenDataBase()

set rs=server.createobject("ADODB.recordset")

sql = "SELECT * FROM yma_customer_recall ORDER BY recall_id"

rs.open sql,conn,1,3

'on error resume next
'conn.Execute sql

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=recall_list.xls"

if rs.eof <> true then
	Response.Write "<table border=1>"
	Response.Write "<tr>"
	Response.Write "<td><strong>ID</strong></td>"
	Response.Write "<td><strong>Dealer</strong></td>"
	Response.Write "<td><strong>Product</strong></td>"
	Response.Write "<td><strong>Qty</strong></td>"
	Response.Write "<td><strong>Customer</strong></td>"		
	Response.Write "<td><strong>Address</strong></td>"
	Response.Write "<td><strong>City</strong></td>"
	Response.Write "<td><strong>State</strong></td>"
	Response.Write "<td><strong>Postcode</strong></td>"	
	Response.Write "<td><strong>Phone</strong></td>"
	Response.Write "<td><strong>Mobile</strong></td>"
	Response.Write "<td><strong>Email</strong></td>"
	Response.Write "<td><strong>Tested by</strong></td>"
	Response.Write "<td><strong>Site Visit</strong></td>"	
	Response.Write "<td><strong>Comments</strong></td>"
	Response.Write "<td><strong>Status</strong></td>"
	Response.Write "</tr>"   
	while not rs.eof
		Response.Write "<tr>"
		Response.Write "<td>" & rs.fields("recall_id") & "</td>"		
		Response.Write "<td>" & rs.fields("dealer") & "</td>"
		Response.Write "<td>" & rs.fields("product") & "</td>"
		Response.Write "<td>" & rs.fields("qty") & "</td>"
		Response.Write "<td>" & rs.fields("customer_name") & "</td>"
		Response.Write "<td>" & rs.fields("customer_address") & "</td>"
		Response.Write "<td>" & rs.fields("customer_city") & "</td>"
		Response.Write "<td>" & rs.fields("customer_state") & "</td>"
		Response.Write "<td>" & rs.fields("customer_postcode") & "</td>"
		Response.Write "<td>" & rs.fields("customer_phone") & "</td>"
		Response.Write "<td>" & rs.fields("customer_mobile") & "</td>"
		Response.Write "<td>" & rs.fields("customer_email") & "</td>"
		Response.Write "<td>" & rs.fields("tested_by") & "</td>"
		if rs.fields("site_visit") = 1 then
			Response.Write "<td>Y</td>"
		else
			Response.Write "<td>N</td>"
		end if		
		Response.Write "<td>" & rs.fields("comments") & "</td>"		
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