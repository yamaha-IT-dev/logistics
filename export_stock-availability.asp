<%@ Language=VBScript %>
<!--#include file="include/connection_base.asp " -->
<%
dim rs
dim strSQL

dim intOnHandsTotal
dim intOnHandsTotalReserved
dim intOnHandsTotalAllocated
dim intOnHandsTotalAvailable
dim intInTransitTotal
dim intInTransitTotalReserved
dim intInTransitTotalAllocated
dim intInTransitTotalAvailable
dim intTotalBackorder
	
intOnHandsTotal = 0
intOnHandsTotalReserved = 0
intOnHandsTotalAllocated = 0
intOnHandsTotalAvailable = 0
intInTransitTotal = 0
intInTransitTotalReserved = 0
intInTransitTotalAllocated = 0
intInTransitTotalAvailable = 0
intTotalBackorder = 0

Call OpenBaseDataBase()

set rs=server.createobject("ADODB.recordset")

	strSQL = "SELECT A5SOSC, A5JZSU, A5JZKH, A5JZHS, A5JZSU - (A5JZKH + A5JZHS) AS AVAILOH, "
	strSQL = strSQL & "	A5MCHS, A5MHKS, A5MHZS, A5MCHS - (A5MHKS + A5MHZS) AS AVAILIT, A5JUSG, A5SOCD, "
	strSQL = strSQL & "	Y3SOMB, Y3STIK, Y1.YINWPR AS RRP, Y2.YINWPR AS S01 "
	strSQL = strSQL & " 	FROM AF5SP "
	strSQL = strSQL & "			INNER JOIN YF3MP ON A5SOSC = Y3SOSC "
	strSQL = strSQL & "			INNER JOIN EF2SP ON A5SOSC = E2SOSC "
	strSQL = strSQL & "			INNER JOIN YFIMP Y1 ON A5SOSC = Y1.YISOSC "	
	strSQL = strSQL & "			INNER JOIN YFIMP Y2 ON A5SOSC = Y2.YISOSC "			
	strSQL = strSQL & "			WHERE A5SOSC LIKE '%" & UCASE(trim(session("stockavail_search"))) & "%' "
	if trim(session("stockavail_warehouse")) = "" then
		strSQL = strSQL & "				AND (A5SOCD LIKE '%3S%' OR A5SOCD LIKE '%3XL%' OR A5SOCD LIKE '%3T%')"
	else
		strSQL = strSQL & "				AND A5SOCD LIKE '%" & UCASE(trim(session("stockavail_warehouse"))) & "%' "
	end if	
	strSQL = strSQL & "				AND A5SKKI <> 'D' "
	strSQL = strSQL & "				AND Y3SKKI <> 'D' "
	strSQL = strSQL & "				AND Y3SOSC NOT LIKE '*%' "
	strSQL = strSQL & "				AND Y3SOSC NOT LIKE '#%' "
	if trim(session("stockavail_department")) = "407" then
		strSQL = strSQL & "				AND LEFT(Y3GREG, 3) = '" & trim(session("stockavail_department")) & "' "
	else
		strSQL = strSQL & "				AND LEFT(Y3GREG, 1) = '" & trim(session("stockavail_department")) & "' "
	end if	
	strSQL = strSQL & "				AND Y3STIK LIKE '%" & UCASE(trim(session("stockavail_item_type"))) & "%' "
	strSQL = strSQL & "				AND E2NGTY = (SELECT E2NGTY FROM EF2SP WHERE E2SOSC = A5SOSC ORDER BY E2NGTY, E2NGTM DESC Fetch First 1 Row Only) "
	strSQL = strSQL & "				AND E2NGTM = (SELECT E2NGTM FROM EF2SP WHERE E2SOSC = A5SOSC ORDER BY E2NGTY, E2NGTM DESC Fetch First 1 Row Only) "
	strSQL = strSQL & "				AND Y1.YISKKI <> 'D' AND Y2.YISKKI <> 'D' "
	strSQL = strSQL & "				AND E2SKKI <> 'D' "
	strSQL = strSQL & "				AND Y1.YIUSPT = 'S50' AND Y2.YIUSPT = 'S01' "
	strSQL = strSQL & "		ORDER BY A5SOSC"

rs.open strSQL,conn,1,3

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=stock-availability.xls"

if rs.eof <> true then
	response.write "<table border=1>"
	response.write "<tr>"
	response.write "<td><strong>Product Code</strong></td>"
	response.write "<td><strong>Description</strong></td>"
	response.write "<td><strong>RRP</strong></td>"
	response.write "<td><strong>S01</strong></td>"
	response.write "<td><strong>Total</strong></td>"
	response.write "<td><strong>Reserved</strong></td>"
	response.write "<td><strong>Allocated</strong></td>"
	response.write "<td><strong>Available</strong></td>"
	response.write "<td><strong>Total</strong></td>"
	response.write "<td><strong>Reserved</strong></td>"
	response.write "<td><strong>Allocated</strong></td>"
	response.write "<td><strong>Available</strong></td>"
	response.write "<td><strong>Backorder</strong></td>"
	response.write "<td><strong>Warehouse</strong></td>"
	response.write "</tr>"
	
	while not rs.eof
		intOnHandsTotal 			= intOnHandsTotal + Cint(rs("A5JZSU"))
		intOnHandsTotalReserved 	= intOnHandsTotalReserved + Cint(rs("A5JZKH"))			
		intOnHandsTotalAllocated 	= intOnHandsTotalAllocated + Cint(rs("A5JZHS"))
		intOnHandsTotalAvailable 	= intOnHandsTotalAvailable + Cint(rs("AVAILOH"))
		intInTransitTotal 			= intInTransitTotal + Cint(rs("A5MCHS"))
		intInTransitTotalReserved 	= intInTransitTotalReserved + Cint(rs("A5MHKS"))
		intInTransitTotalAllocated 	= intInTransitTotalAllocated + Cint(rs("A5MHZS"))
		intInTransitTotalAvailable 	= intInTransitTotalAvailable + Cint(rs("AVAILIT"))
		intTotalBackorder 			= intTotalBackorder + Cint(rs("A5JUSG"))
		
		response.write "<tr>"
		response.write "<td>" & rs.fields("A5SOSC") & "</td>"
		response.write "<td>" & rs.fields("Y3SOMB") & "</td>"
		response.write "<td>" & FormatNumber(rs.fields("RRP")) & "</td>"
		response.write "<td>" & FormatNumber(rs.fields("S01")) & "</td>"
		response.write "<td>" & rs.fields("A5JZSU") & "</td>"
		response.write "<td>" & rs.fields("A5JZKH") & "</td>"
		response.write "<td>" & rs.fields("A5JZHS") & "</td>"
		response.write "<td>" & rs.fields("AVAILOH") & "</td>"
		response.write "<td>" & rs.fields("A5MCHS") & "</td>"
		response.write "<td>" & rs.fields("A5MHKS") & "</td>"
		response.write "<td>" & rs.fields("A5MHZS") & "</td>"
		response.write "<td>" & rs.fields("AVAILIT") & "</td>"
		response.write "<td>" & rs.fields("A5JUSG") & "</td>"
		response.write "<td>" & rs.fields("A5SOCD") & "</td>"
		response.write "</tr>"
		rs.movenext
	wend
	response.write "<tr>"
	response.write "<td><b>Grand Total:</b></td>"
	response.write "<td></td>"
	response.write "<td></td>"
	response.write "<td></td>"
	response.write "<td>" & intOnHandsTotal & "</td>"
	response.write "<td>" & intOnHandsTotalReserved & "</td>"
	response.write "<td>" & intOnHandsTotalAllocated & "</td>"
	response.write "<td>" & intOnHandsTotalAvailable & "</td>"
	response.write "<td>" & intInTransitTotal & "</td>"
	response.write "<td>" & intInTransitTotalReserved & "</td>"
	response.write "<td>" & intInTransitTotalAllocated & "</td>"
	response.write "<td>" & intInTransitTotalAvailable & "</td>"
	response.write "<td>" & intTotalBackorder & "</td>"
	response.write "<td></td>"
	response.write "</tr>"
	response.write "</table>"
end if

Call CloseBaseDataBase()
%>