<%@ Language=VBScript %>
<!--#include file="include/connection_it.asp " -->
<%
dim rs
dim strSQL

Call OpenDataBase()

set rs=server.createobject("ADODB.recordset")

strSQL = "SELECT * FROM logistic_pack "
	strSQL = strSQL & "	WHERE  (packEmail LIKE '%" & session("pack_search") & "%' "
	strSQL = strSQL & "			OR packName LIKE '%" & session("pack_search") & "%' "
	strSQL = strSQL & "			OR packComments LIKE '%" & session("pack_search") & "%') "
	strSQL = strSQL & "		AND packstatus LIKE '%" & session("pack_status") & "%' "
	strSQL = strSQL & "	ORDER BY packDateCreated DESC"

rs.open strSQL,conn,1,3

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=pack-request_list.xls"

if rs.eof <> true then
	Response.Write "<table border=1>"
	Response.Write "<tr>"
	Response.Write "<td><strong>Created by</strong></td>"
	Response.Write "<td><strong>Date Created</strong></td>"
	Response.Write "<td><strong>Name</strong></td>"
	Response.Write "<td><strong>Qty</strong></td>"
	Response.Write "<td><strong>Priority</strong></td>"
	Response.Write "<td><strong>Logistics Confirm</strong></td>"
	Response.Write "<td><strong>Warehouse Confirm</strong></td>"	
	Response.Write "<td><strong>Status</strong></td>"
	Response.Write "<td><strong>Last Modified by</strong></td>"
	Response.Write "<td><strong>Last Modified Date</strong></td>"	
	Response.Write "</tr>"   
	while not rs.eof
		Response.Write "<tr>"
		Response.Write "<td>" & rs.fields("packEmail") & "</td>"
		Response.Write "<td>" & rs.fields("packDateCreated") & "</td>"
		Response.Write "<td>" & rs.fields("packName") & "</td>"
		Response.Write "<td>" & rs.fields("packQty") & "</td>"
		Response.Write "<td>" & rs.fields("packPriority") & "</td>"
		if rs.fields("packLogistics") = 1 then
			Response.Write "<td>Y</td>"
		else
			Response.Write "<td>N</td>"
		end if
		if rs.fields("packWarehouse") = 1 then
			Response.Write "<td>Y</td>"
		else
			Response.Write "<td>N</td>"
		end if
		
		if rs.fields("packStatus") = 1 then
			Response.Write "<td>Open</td>"
		else
			Response.Write "<td>Completed</td>"
		end if
		
		Response.Write "<td>" & rs.fields("packModifiedBy") & "</td>"
		Response.Write "<td>" & rs.fields("packDateModified") & "</td>"				
		Response.Write "</tr>"
		rs.movenext
	wend
	Response.Write "</table>"
end if

Call CloseDataBase()
%>