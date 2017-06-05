<%@ Language=VBScript %>
<!--#include file="include/connection_it.asp " -->
<%
dim rs
dim strSQL

dim strSearch
strSearch 	= Trim(Request("search"))

dim strType
strType 	= Trim(Request("type"))

dim strYear
strYear 	= Trim(Request("year"))

dim intStatus
intStatus 	= Trim(Request("status"))

dim strSort
strSort 	= Trim(Request("sort"))

Call OpenDataBase()

set rs=server.createobject("ADODB.recordset")

strSQL = "SELECT * FROM yma_damage "
strSQL = strSQL & "	WHERE "
if strYear <> "" then
	strSQL = strSQL & " 	YEAR(date_created) = '" & trim(strYear) & "' AND "
end if
strSQL = strSQL & "		damage_type LIKE '%" & strType & "%' "
strSQL = strSQL & "		AND (damage_item LIKE '%" & strSearch & "%' "
strSQL = strSQL & "			OR damage_serial_no LIKE '%" & strSearch & "%') "
strSQL = strSQL & "		AND status LIKE '%" & intStatus & "%' "
strSQL = strSQL & "	ORDER BY " & strSort

rs.open strSQL,conn,1,3

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=damage-"& strYear & "_list.xls"

if rs.eof <> true then
	Response.Write "<table border=1>"
	Response.Write "<tr>"
	Response.Write "<td><strong>ID</strong></td>"
	Response.Write "<td><strong>Damaged Item</strong></td>"
	Response.Write "<td><strong>LIC</strong></td>"
	Response.Write "<td><strong>Serial No</strong></td>"
	Response.Write "<td><strong>Damage Type</strong></td>"	
	Response.Write "<td><strong>Course of Damage</strong></td>"
	Response.Write "<td><strong>Sent to Excel</strong></td>"
	Response.Write "<td><strong>Sent to Excel Date</strong></td>"
	Response.Write "<td><strong>Comments</strong></td>"
	Response.Write "<td><strong>Date Created</strong></td>"
	Response.Write "<td><strong>Status</strong></td>"
	Response.Write "</tr>"   
	while not rs.eof
		Response.Write "<tr>"
		Response.Write "<td>" & rs.fields("damage_id") & "</td>"
		Response.Write "<td>" & rs.fields("damage_item") & "</td>"
		Response.Write "<td>" & rs.fields("lic") & "</td>"
		Response.Write "<td>" & rs.fields("damage_serial_no") & "</td>"
		Response.Write "<td>" & rs.fields("damage_type") & "</td>"
		Response.Write "<td>" & rs.fields("course_damage") & "</td>"
		'Response.Write "<td>" & rs.fields("sent_excel") & "</td>"
		if rs.fields("sent_excel") = 1 then
			Response.Write "<td>Y</td>"
		else
			Response.Write "<td>N</td>"
		end if
		Response.Write "<td>" & rs.fields("sent_excel_date") & "</td>"
		Response.Write "<td>" & rs.fields("damage_comments") & "</td>"
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