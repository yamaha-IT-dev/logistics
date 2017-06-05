<%@ Language=VBScript %>
<!--#include file="include/connection_it.asp " -->
<%
dim rs
dim strSQL
dim strDept
strDept = Trim(Request("cboDepartment"))

Call OpenDataBase()

set rs=server.createobject("ADODB.recordset")

strSQL = "SELECT * FROM yma_roadshow_2014 "
	strSQL = strSQL & "	WHERE category LIKE '%" & session("roadshow_Category") & "%' "
	strSQL = strSQL & "		AND department LIKE '%" & session("roadshow_ItemDepartment") & "%' "
	strSQL = strSQL & "		AND sku_type LIKE '%" & session("roadshow_SkuType") & "%' "
	strSQL = strSQL & "		AND transit LIKE '%" & session("roadshow_Transit") & "%' "
	strSQL = strSQL & "		AND return_to LIKE '%" & session("roadshow_ReturnTo") & "%' "
	strSQL = strSQL & "		AND origin LIKE '%" & session("roadshow_ItemOrigin") & "%' "
	strSQL = strSQL & "		AND created_by LIKE '%" & session("roadshow_ItemOwner") & "%' "
	strSQL = strSQL & "		AND (product_code LIKE '%" & session("roadshow_Search") & "%' "
	strSQL = strSQL & "			OR description LIKE '%" & session("roadshow_Search") & "%' "
	strSQL = strSQL & "			OR pallet_no LIKE '%" & session("roadshow_Search") & "%' "
	strSQL = strSQL & "			OR loading_sequence LIKE '%" & session("roadshow_Search") & "%') "
	strSQL = strSQL & "	ORDER BY item_id"

rs.open strSQL,conn,1,3

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=roadshow-2014-list.xls"

if rs.eof <> true then
	Response.Write "<table border=1>"
	Response.Write "<tr>"
	Response.Write "<td><strong>ID</strong></td>"
	Response.Write "<td><strong>Owner</strong></td>"
	Response.Write "<td><strong>Dept</strong></td>"	
	Response.Write "<td><strong>Item Group</strong></td>"
	Response.Write "<td><strong>Category</strong></td>"	
	Response.Write "<td><strong>Item</strong></td>"
	Response.Write "<td><strong>Description</strong></td>"		
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
	Response.Write "<td><strong>Display</strong></td>"
	Response.Write "<td><strong>How Displayed</strong></td>"
	Response.Write "<td><strong>F and B available</strong></td>"	
	Response.Write "<td><strong>Invoice No</strong></td>"	
	Response.Write "<td><strong>Pallet No</strong></td>"	
	Response.Write "<td><strong>Loading No</strong></td>"		
	Response.Write "<td><strong>Status</strong></td>"
	Response.Write "</tr>"   
	while not rs.eof
		Response.Write "<tr>"
		Response.Write "<td>" & rs.fields("item_id") & "</td>"
		Response.Write "<td>" & rs.fields("owner") & "</td>"
		Response.Write "<td>" & rs.fields("department") & "</td>"		
		Response.Write "<td>" & rs.fields("item_group") & "</td>"
		Response.Write "<td>" & rs.fields("category") & "</td>"
		Response.Write "<td>" & rs.fields("product_code") & "</td>"
		Response.Write "<td>" & rs.fields("description") & "</td>"		
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
		Response.Write "<td>" & rs.fields("displayed") & "</td>"
		Response.Write "<td>" & rs.fields("how_displayed") & "</td>"
		Response.Write "<td>" & rs.fields("fb_completed") & "</td>"
		Response.Write "<td>" & rs.fields("invoice_no") & "</td>"
		Response.Write "<td>" & rs.fields("pallet_no") & "</td>"
		Response.Write "<td>" & rs.fields("loading_sequence") & "</td>"				
		if rs.fields("status") = 1 then
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