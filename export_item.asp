<%@ Language=VBScript %>
<!--#include file="include/connection_it.asp " -->
<%
dim rs
dim sql
dim strDept
strDept = Trim(Request("cboDepartment"))

Call OpenDataBase()

set rs=server.createobject("ADODB.recordset")

sql = "SELECT * FROM yma_item ORDER BY item_id"

rs.open sql,conn,1,3

'on error resume next
'conn.Execute sql

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=roadshow_list.xls"

if rs.eof <> true then
	Response.Write "<table border=1>"
	Response.Write "<tr>"
	Response.Write "<td><strong>Owner</strong></td>"
	Response.Write "<td><strong>Item Code</strong></td>"
	Response.Write "<td><strong>Item Group</strong></td>"
	Response.Write "<td><strong>Description</strong></td>"
	Response.Write "<td><strong>Category</strong></td>"	
	Response.Write "<td><strong>Dept</strong></td>"
	Response.Write "<td><strong>RRP</strong></td>"
	Response.Write "<td><strong>SKU Type</strong></td>"
	Response.Write "<td><strong>Prototype</strong></td>"
	Response.Write "<td><strong>Qty</strong></td>"
	Response.Write "<td><strong>Packaging</strong></td>"
	Response.Write "<td><strong>Source</strong></td>"
	Response.Write "<td><strong>Origin</strong></td>"
	Response.Write "<td><strong>Available</strong></td>"
	Response.Write "<td><strong>In-transit</strong></td>"	
	Response.Write "<td><strong>Type</strong></td>"
	Response.Write "<td><strong>Pre-sold</strong></td>"
	Response.Write "<td><strong>Aftershow Return to</strong></td>"
	Response.Write "<td><strong>Invoice No</strong></td>"	
	Response.Write "<td><strong>Pallet No</strong></td>"	
	Response.Write "<td><strong>Loading No</strong></td>"		
	Response.Write "<td><strong>Status</strong></td>"
	Response.Write "</tr>"   
	while not rs.eof
		Response.Write "<tr>"
		Response.Write "<td>" & rs.fields("owner") & "</td>"
		Response.Write "<td>" & rs.fields("product_code") & "</td>"
		Response.Write "<td>" & rs.fields("item_group") & "</td>"
		Response.Write "<td>" & rs.fields("description") & "</td>"
		Response.Write "<td>" & rs.fields("category") & "</td>"
		Response.Write "<td>" & rs.fields("department") & "</td>"
		Response.Write "<td>" & rs.fields("rrp") & "</td>"
		Response.Write "<td>" & rs.fields("sku_type") & "</td>"
		Response.Write "<td>" & rs.fields("prototype") & "</td>"
		Response.Write "<td>" & rs.fields("quantity") & "</td>"
		Response.Write "<td>" & rs.fields("packaging") & "</td>"
		Response.Write "<td>" & rs.fields("source") & "</td>"
		Response.Write "<td>" & rs.fields("origin") & "</td>"
		Response.Write "<td>" & rs.fields("available") & "</td>"
		Response.Write "<td>" & rs.fields("transit") & "</td>"
		Response.Write "<td>" & rs.fields("type") & "</td>"
		Response.Write "<td>" & rs.fields("pre_sold") & "</td>"		
		Response.Write "<td>" & rs.fields("return_to") & "</td>"		
		Response.Write "<td>" & rs.fields("invoice_no") & "</td>"
		Response.Write "<td>" & rs.fields("pallet_no") & "</td>"
		Response.Write "<td>" & rs.fields("loading_sequence") & "</td>"				
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