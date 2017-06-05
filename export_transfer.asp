<%@ Language=VBScript %>
<!--#include file="include/connection_it.asp " -->
<%
dim rs
dim sql

dim strSearch
strSearch 	= Trim(Request("search"))

dim strDept
strDept = Trim(Request("cboDepartment"))

dim intStatus
intStatus 	= Trim(Request("status"))

Call OpenDataBase()

set rs=server.createobject("ADODB.recordset")

sql = "SELECT * FROM yma_transfer WHERE (product_1 LIKE '%" & strSearch & "%' OR product_2 LIKE '%" & strSearch & "%' OR product_3 LIKE '%" & strSearch & "%' OR product_4 LIKE '%" & strSearch & "%' OR product_5 LIKE '%" & strSearch & "%' OR product_6 LIKE '%" & strSearch & "%' OR product_7 LIKE '%" & strSearch & "%' OR product_8 LIKE '%" & strSearch & "%' OR product_9 LIKE '%" & strSearch & "%' OR product_10 LIKE '%" & strSearch & "%' OR product_11 LIKE '%" & strSearch & "%' OR product_12 LIKE '%" & strSearch & "%' OR product_13 LIKE '%" & strSearch & "%' OR product_14 LIKE '%" & strSearch & "%' OR product_15 LIKE '%" & strSearch & "%' OR product_16 LIKE '%" & strSearch & "%' OR product_17 LIKE '%" & strSearch & "%' OR product_18 LIKE '%" & strSearch & "%' OR product_19 LIKE '%" & strSearch & "%' OR product_20 LIKE '%" & strSearch & "%') AND status LIKE '%" & intStatus & "%'"

rs.open sql,conn,1,3

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=transfer_list.xls"

if rs.eof <> true then
	Response.Write "<table border=1>"
	Response.Write "<tr>"
	Response.Write "<td><strong>ID</strong></td>"
	Response.Write "<td><strong>Requested By</strong></td>"
	Response.Write "<td><strong>From - To</strong></td>"
	Response.Write "<td><strong>Products</strong></td>"
	Response.Write "<td><strong>Transfer Date</strong></td>"	
	Response.Write "<td><strong>Date Received</strong></td>"
	Response.Write "<td><strong>Transfer Comments</strong></td>"
	Response.Write "<td><strong>Picked Up?</strong></td>"
	Response.Write "<td><strong>Received?</strong></td>"
	Response.Write "<td><strong>Invoice No</strong></td>"
	Response.Write "<td><strong>Date Created</strong></td>"
	Response.Write "<td><strong>Status</strong></td>"
	Response.Write "</tr>"   
	while not rs.eof
		Response.Write "<tr>"
		Response.Write "<td>" & rs.fields("id") & "</td>"
		Response.Write "<td>" & rs.fields("created_by") & "</td>"
		Response.Write "<td>" & rs.fields("warehouse") & "</td>"
		Response.Write "<td>" & rs.fields("product_1") & " "
		if rs.fields("product_2") <> "" then
				Response.Write ", " & rs.fields("product_2")
			end if
			if rs.fields("product_3") <> "" then
				Response.Write ", " & rs.fields("product_3")
			end if
			if rs.fields("product_4") <> "" then
				Response.Write ", " & rs.fields("product_4")
			end if
			if rs.fields("product_5") <> "" then
				Response.Write ", " & rs.fields("product_5")
			end if	
			if rs.fields("product_6") <> "" then
				Response.Write ", " & rs.fields("product_6")
			end if
			if rs.fields("product_7") <> "" then
				Response.Write ", " & rs.fields("product_7")
			end if		
			if rs.fields("product_8") <> "" then
				Response.Write ", " & rs.fields("product_8")
			end if
			if rs.fields("product_9") <> "" then
				Response.Write ", " & rs.fields("product_9")
			end if
			if rs.fields("product_10") <> "" then
				Response.Write ", " & rs.fields("product_10")
			end if
			if rs.fields("product_11") <> "" then
				Response.Write ", " & rs.fields("product_11")
			end if
			if rs.fields("product_12") <> "" then
				Response.Write ", " & rs.fields("product_12")
			end if
			if rs.fields("product_13") <> "" then
				Response.Write ", " & rs.fields("product_13")
			end if
			if rs.fields("product_14") <> "" then
				Response.Write ", " & rs.fields("product_14")
			end if
			if rs.fields("product_15") <> "" then
				Response.Write ", " & rs.fields("product_15")
			end if	
			if rs.fields("product_16") <> "" then
				Response.Write ", " & rs.fields("product_16")
			end if
			if rs.fields("product_17") <> "" then
				Response.Write ", " & rs.fields("product_17")
			end if		
			if rs.fields("product_18") <> "" then
				Response.Write ", " & rs.fields("product_18")
			end if
			if rs.fields("product_19") <> "" then
				Response.Write ", " & rs.fields("product_19")
			end if
			if rs.fields("product_20") <> "" then
				Response.Write ", " & rs.fields("product_20")
			end if
			Response.Write "</td>"
		Response.Write "<td>" & rs.fields("transfer_date") & "</td>"
		Response.Write "<td>" & rs.fields("date_received") & "</td>"
		Response.Write "<td>" & rs.fields("transfer_comments") & "</td>"
		if rs.fields("pickup") = 1 then
			Response.Write "<td>Y</td>"
		else
			Response.Write "<td>N</td>"
		end if
		if rs.fields("received") = 1 then
			Response.Write "<td>Y</td>"
		else
			Response.Write "<td>N</td>"
		end if
		Response.Write "<td>" & rs.fields("invoice_no") & "</td>"
		Response.Write "<td>" & rs.fields("date_created") & "</td>"		
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