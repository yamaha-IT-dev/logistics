<%@ Language=VBScript %>
<!--#include file="include/connection_it.asp " -->
<%
dim rs
dim strSQL
dim strDept
strDept = Trim(Request("cboDepartment"))

Call OpenDataBase()

set rs=server.createobject("ADODB.recordset")

strSQL = "SELECT * FROM tbl_focus "
	strSQL = strSQL & "	WHERE product_code LIKE '%" & session("strSearch") & "%' "
	strSQL = strSQL & "			OR stock_situation LIKE '%" & session("strSearch") & "%' "
	strSQL = strSQL & "			OR display LIKE '%" & session("strSearch") & "%' "
	strSQL = strSQL & "			OR instruction LIKE '%" & session("strSearch") & "%' "
	strSQL = strSQL & "	ORDER BY id"

rs.open strSQL,conn,1,3

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=focus-list.xls"

if rs.eof <> true then
	Response.Write "<table border=1>"
	Response.Write "<tr>"
	Response.Write "<td><strong>ID</strong></td>"
	Response.Write "<td><strong>Owner</strong></td>"
	Response.Write "<td><strong>Type</strong></td>"
	Response.Write "<td><strong>Qty</strong></td>"
	Response.Write "<td><strong>Product Code</strong></td>"
	Response.Write "<td><strong>Location</strong></td>"	
	Response.Write "<td><strong>Stock Situation</strong></td>"
	Response.Write "<td><strong>Loan Account</strong></td>"		
	Response.Write "<td><strong>Display</strong></td>"
	Response.Write "<td><strong>For Sale</strong></td>"
	Response.Write "<td><strong>Instruction</strong></td>"	
	Response.Write "<td><strong>Notes</strong></td>"
	Response.Write "</tr>"   
	while not rs.eof
		Response.Write "<tr>"
		Response.Write "<td>" & rs.fields("id") & "</td>"
		Response.Write "<td>" & rs.fields("owner") & "</td>"
		Response.Write "<td>" & rs.fields("type") & "</td>"		
		Response.Write "<td>" & rs.fields("quantity") & "</td>"
		Response.Write "<td>" & rs.fields("product_code") & "</td>"
		Response.Write "<td>" & rs.fields("location") & "</td>"
		Response.Write "<td>" & rs.fields("stock_situation") & "</td>"		
		Response.Write "<td>" & rs.fields("loan_account") & "</td>"
		Response.Write "<td>" & rs.fields("display") & "</td>"
		Response.Write "<td>" & rs.fields("available_for_sale") & "</td>"
		Response.Write "<td>" & rs.fields("instruction") & "</td>"
		Response.Write "<td>" & rs.fields("notes") & "</td>"		
		Response.Write "</tr>"
		rs.movenext
	wend
	Response.Write "</table>"
end if

Call CloseDataBase()
%>