<%@ Language=VBScript %>
<!--#include file="include/connection_it.asp " -->
<%
dim rs
dim strSQL

Call OpenDataBase()

set rs=server.createobject("ADODB.recordset")

	strSQL = "SELECT * FROM yma_shipment "
	strSQL = strSQL & " WHERE (supplier_invoice_no LIKE '%" & session("shipment_search") & "%' "
	strSQL = strSQL & "			OR container_no LIKE '%" & session("shipment_search") & "%' "
	strSQL = strSQL & "			OR EFT LIKE '%" & session("shipment_search") & "%' "
	strSQL = strSQL & "			OR vessel_name LIKE '%" & session("shipment_search") & "%') "
	strSQL = strSQL & "		AND department LIKE '%" & session("shipment_department") & "%' "
	strSQL = strSQL & "		AND country_origin LIKE '%" & session("shipment_country") & "%' "
	strSQL = strSQL & "		AND warehouse LIKE '%" & session("shipment_warehouse") & "%' "
	strSQL = strSQL & "		AND fta_status LIKE '%" & session("shipment_fta") & "%' "
	strSQL = strSQL & "		AND status = 1 "
	strSQL = strSQL & "	ORDER BY eta_unpacked DESC, eta_discharged DESC, supplier_invoice_no"

rs.open strSQL,conn,1,3

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=shipment_list.xls"

if rs.eof <> true then
	Response.Write "<table border=1>"
	Response.Write "<tr>"
	Response.Write "<td><strong>Container No</strong></td>"
	Response.Write "<td><strong>Supplier Invoice No</strong></td>"
	Response.Write "<td><strong>Dept</strong></td>"
	Response.Write "<td><strong>Custom Cleared</strong></td>"	
	Response.Write "<td><strong>Fumigation</strong></td>"
	Response.Write "<td><strong>FF</strong></td>"
	Response.Write "<td><strong>EFT</strong></td>"
	Response.Write "<td><strong>All Docs</strong></td>"
	Response.Write "<td><strong>Commodity</strong></td>"	
	Response.Write "<td><strong>Port</strong></td>"
	Response.Write "<td><strong>Country</strong></td>"
	Response.Write "<td><strong>Vessel Name</strong></td>"
	Response.Write "<td><strong>Voyage</strong></td>"
	Response.Write "<td><strong>Warehouse</strong></td>"
	Response.Write "<td><strong>No of Cartons</strong></td>"
	Response.Write "<td><strong>Shipment</strong></td>"
	Response.Write "<td><strong>Melb ETA</strong></td>"
	Response.Write "<td><strong>Melb ADT</strong></td>"
	Response.Write "<td><strong>Container ETA</strong></td>"
	Response.Write "<td><strong>Container ETA</strong></td>"
	Response.Write "<td><strong>Unpack ETA</strong></td>"
	Response.Write "<td><strong>Unpack Time</strong></td>"
	Response.Write "<td><strong>TEU</strong></td>"
	Response.Write "<td><strong>Paperwork</strong></td>"
	Response.Write "<td><strong>Delivery</strong></td>"
	Response.Write "</tr>"   
	while not rs.eof
		Response.Write "<tr>"
		Response.Write "<td>" & rs.fields("container_no") & "</td>"
		Response.Write "<td><a href=""file:\\YAMMAS22\shipment\" & rs("supplier_invoice_no") & """ target=""_blank"">" & rs.fields("supplier_invoice_no") & "</a></td>"
		Response.Write "<td>" & rs.fields("department") & "</td>"
		Response.Write "<td>" & rs.fields("custom_cleared") & "</td>"
		Response.Write "<td>" & rs.fields("fumigation") & "</td>"
		Response.Write "<td>" & rs.fields("FF") & "</td>"
		Response.Write "<td>" & rs.fields("EFT") & "</td>"
		Response.Write "<td>" & rs.fields("all_documents") & "</td>"
		Response.Write "<td>" & rs.fields("commodity") & "</td>"
		Response.Write "<td>" & rs.fields("port_origin") & "</td>"
		Response.Write "<td>" & rs.fields("country_origin") & "</td>"
		Response.Write "<td>" & rs.fields("vessel_name") & "</td>"
		Response.Write "<td>" & rs.fields("voyage") & "</td>"
		Response.Write "<td>" & rs.fields("warehouse") & "</td>"
		Response.Write "<td>" & rs.fields("cartons") & "</td>"
		Response.Write "<td>" & rs.fields("date_shipment") & "</td>"
		Response.Write "<td>" & rs.fields("eta_discharged") & "</td>"
		Response.Write "<td>" & rs.fields("melb_time") & "</td>"
		Response.Write "<td>" & rs.fields("eta_availability") & "</td>"
		Response.Write "<td>" & rs.fields("melb_eta_time") & "</td>"
		Response.Write "<td>" & rs.fields("eta_unpacked") & "</td>"
		Response.Write "<td>" & rs.fields("unpack_time") & "</td>"
		Response.Write "<td>" & rs.fields("teu") & "</td>"
		Response.Write "<td>" & rs.fields("paperwork") & "</td>"
		Response.Write "<td>" & rs.fields("delivery_type") & "</td>"
		Response.Write "</tr>"
		rs.movenext
	wend
	Response.Write "</table>"
end if

Call CloseDataBase()
%>