<!--#include file="include/connection_it.asp " -->
<%
dim rs
dim strSQL

Call OpenDataBase()

set rs = server.createobject("ADODB.recordset")

	strSQL = "SELECT DISTINCT BTUSNO, BTKASC, BTOPEC, BTHYNO, BTURKC + BTHSRC AS dealer_code, BTAHYY, BTAHYM, BTAHYD, " 
	'strSQL = strSQL & "  cast(BTAHYY as varchar(4)) + right('00'+cast(BTAHYM as varchar(2)),2) + right('00'+cast(BTAHYD as varchar(2)),2) return_date, "
	strSQL = strSQL & "  BTRTST, BTUSGR, BTATT1, BTSCNO, BTGICM, BTNICM, "
	strSQL = strSQL & "  YMOPNM, Y1JSO1, Y1JSO3, Y1UBNB, Y1KOM1, Y1KNCI, Y1TELN, gra_connote, "
	strSQL = strSQL & "  	CASE "
	strSQL = strSQL & "			WHEN deduction_gra_no IS NULL THEN 'no' "
	strSQL = strSQL & "			ELSE 'yes'"
	strSQL = strSQL & "  	END AS 'deduction' "
	strSQL = strSQL & "  FROM OPENQUERY(AS400, 'SELECT BTUSNO, BTKASC, BTOPEC, BTHYNO, BTURKC, BTHSRC, BTAHYY, BTAHYM, BTAHYD, BTAHYY, BTAHYM, BTAHYD, BTRTST, BTUSGR, BTATT1, BTSCNO, BTGICM, BTNICM, BTSKKI FROM BFTEP') "
	strSQL = strSQL & "  	INNER JOIN OPENQUERY(AS400, 'SELECT BUHYNO, BUSOSC, BUCLMN, BUSIBN FROM BFUEP') ON BTHYNO = BUHYNO "
	strSQL = strSQL & " 		LEFT JOIN OPENQUERY(AS400, 'SELECT Y1KOM1, Y1JSO1, Y1JSO3, Y1KNCI, Y1UBNB, Y1TELN, Y1KOKC FROM YF1MP WHERE Y1SKKI <> ''D''') ON BTURKC + BTHSRC = Y1KOKC "
	strSQL = strSQL & " 		LEFT JOIN (SELECT YMOPEC, MAX(('20' + RIGHT(RTRIM(YMKOYY),2)) * 10000 + SUBSTRING(CONVERT(VARCHAR(6), YMKOYY), 3,2) * 100 + LEFT(YMKOYY,2)) AS MOD_DATE"
	strSQL = strSQL & " 		FROM AS400.S1027CFG.YGZFLIB.YFMMP WHERE RIGHT(RTRIM(YMKOYY),2) < 97"
	strSQL = strSQL & " 			GROUP BY YMOPEC"
	strSQL = strSQL & " 			) AS OP ON BTOPEC = OP.YMOPEC"
	strSQL = strSQL & " 	INNER JOIN (SELECT YMOPEC, YMOPNM, YMPMID, (('20' + RIGHT(RTRIM(YMKOYY),2)) * 10000 + SUBSTRING(CONVERT(VARCHAR(6), YMKOYY), 3,2) * 100 + LEFT(YMKOYY,2)) AS MOD_DATE"
	strSQL = strSQL & " 	FROM AS400.S1027CFG.YGZFLIB.YFMMP"
	strSQL = strSQL & " 	) AS OP_NAME ON OP.YMOPEC = OP_NAME.YMOPEC AND OP.MOD_DATE = OP_NAME.MOD_DATE"	
	strSQL = strSQL & " 		LEFT JOIN tbl_deductions on BTHYNO = deduction_gra_no "
	strSQL = strSQL & " 		LEFT JOIN yma_gra_status on BTHYNO = gra_no "
	strSQL = strSQL & " WHERE (BTHYNO LIKE '%" & session("gra_search") & "%' "
	strSQL = strSQL & " 		OR BTURKC LIKE '%" & UCASE(session("gra_search")) & "%' "
	strSQL = strSQL & " 		OR BTATT1 LIKE '%" & UCASE(session("gra_search")) & "%' "
	strSQL = strSQL & " 		OR YMOPNM LIKE '%" & UCASE(session("gra_search")) & "%' "
	strSQL = strSQL & " 		OR Y1KOM1 LIKE '%" & UCASE(session("gra_search")) & "%' "
	strSQL = strSQL & "			OR BUSIBN LIKE '%" & UCASE(session("gra_search")) & "%' "
	strSQL = strSQL & "			OR BUCLMN LIKE '%" & UCASE(session("gra_search")) & "%' "
	strSQL = strSQL & "			OR gra_connote LIKE '%" & UCASE(session("gra_search")) & "%' "
	strSQL = strSQL & "			OR BUSOSC LIKE '%" & UCASE(session("gra_search")) & "%')"
	strSQL = strSQL & " 	AND BTSKKI <> 'D' "
	if trim(session("gra_search_month")) <> "" then
		strSQL = strSQL & " 	AND BTAHYM = '" & session("gra_search_month") & "' "
	end if
	strSQL = strSQL & " 	AND BTAHYY = '" & session("gra_search_year") & "' "
	strSQL = strSQL & " 	AND YMOPNM LIKE '%" & UCASE(session("gra_search_operator")) & "%' "
	strSQL = strSQL & " 	AND BTSCNO LIKE '%" & session("gra_search_warehouse") & "%' "
	strSQL = strSQL & "		AND BTRTST LIKE '%" & session("gra_search_status") & "%' "
	strSQL = strSQL & "		AND (YMPMID like 'LOG%' OR YMPMID like 'INT%' OR YMPMID like 'SERV%' OR YMPMID like 'CREDIT%')"
	strSQL = strSQL & " 	ORDER BY BTHYNO DESC "	

rs.open strSQL,conn,1,3

Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=all-gra.xls"

if rs.eof <> true then
	Response.Write "<table border=1>"
	Response.Write "<tr>"
	Response.Write "<td nowrap><strong>Created by</strong></td>"
	Response.Write "<td nowrap><strong>GRA no</strong></td>"
	Response.Write "<td nowrap><strong>Dealer Code</strong></td>"
	Response.Write "<td nowrap><strong>Dealer</strong></td>"
	Response.Write "<td nowrap><strong>State</strong></td>"	
	Response.Write "<td nowrap><strong>Comments</strong></td>"	
	Response.Write "<td nowrap><strong>Phone</strong></td>"
	Response.Write "<td nowrap><strong>Plan Return Date</strong></td>"
	Response.Write "<td nowrap><strong>Return Status</strong></td>"
	Response.Write "<td nowrap><strong>Carrier Code</strong></td>"
	Response.Write "<td nowrap><strong>Con-note</strong></td>"	
	Response.Write "</tr>"
	while not rs.eof
		Response.Write "<tr>"
		Response.Write "<td nowrap>" & rs.fields("YMOPNM") & " (" & rs.fields("BTOPEC") & ")</td>"
		Response.Write "<td nowrap>" & rs.fields("BTHYNO") & "</td>"
		Response.Write "<td nowrap>" & rs.fields("dealer_code") & "</td>"
		Response.Write "<td nowrap>" & rs.fields("Y1KOM1") & "</td>"
		Response.Write "<td nowrap>" & rs.fields("Y1KNCI") & "</td>"
		Response.Write "<td nowrap>" & rs.fields("BTGICM") & " " & rs.fields("BTNICM") & "</td>"
		Response.Write "<td nowrap>" & rs.fields("Y1TELN") & "</td>"
		Response.Write "<td nowrap>" & rs.fields("BTAHYD") & "/" & rs.fields("BTAHYM") & "/" & rs.fields("BTAHYY") & "</td>"
		select case rs.fields("BTRTST") 
			case 0
				Response.Write "<td>Not received</td>"
			case 1
				Response.Write "<td>Received not credited</td>"
			case 2
				Response.Write "<td>Credited</td>"			
			case else
				Response.Write "<td>" & rs.fields("BTRTST") & "</td>"
		end select
		Response.Write "<td nowrap>" & rs.fields("BTUSGR") & "</td>"
		Response.Write "<td nowrap>" & rs.fields("gra_connote") & "</td>"
		Response.Write "<td nowrap>" & rs.fields("BTSCNO") & "</td>"		
		Response.Write "</tr>"				
		rs.movenext
	wend
	Response.Write "</table>"
end if

Call CloseDataBase()
%>