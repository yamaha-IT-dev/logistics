<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<!--#include file="class/clsGoodsReturnReport.asp " -->
<!--#include file="class/clsPallet.asp " -->
<% strSection = "gra" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>View Pallet - with GRA Report</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
</head>
<body>
<%
sub displayGraReports
	dim iRecordCount
	iRecordCount = 0

	dim intTotalLabourCount
	intTotalLabourCount = 0

	dim intTotalPartsCount
	intTotalPartsCount = 0

	dim intTotalCount
	intTotalCount = 0

    dim strReportSortBy
	dim strReportSortItem
    dim strReportSearch
    dim strSQL
	dim strReportSort
	dim strStatus

	dim strPageResultNumber
	dim strRecordPerPage
	dim intRecordCount
	dim strModifiedDate

	dim strTodayDate
	strTodayDate = FormatDateTime(Date())

	if session("gra_report_sort") = "" then
		session("gra_report_sort") = "gra_no"
	end if

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation 	= 3	'adUseClient
    rs.CursorType 		= 3	'adOpenStatic
	rs.PageSize 		= 1000

    strSQL = "SELECT GR.*, E2SOSC, E2NGTY, E2NGTM, (E2ihtn+(e2ihtn*e2kzrt/100)+(e2ihtn*e2skkr/100)) AS LIC "
    strSql = strSQL & " FROM yma_gra_report GR "
    strSQL = strSQL & " LEFT JOIN AS400.S1027CFG.YGZFLIB.EF2SP ON item = E2SOSC AND E2NGTY = YEAR(getdate()) AND E2NGTM = MONTH(getdate()) "
    strSQL = strSQL & " WHERE pallet_no = '" & Request("pallet_no") & "' "
    strSQL = strSQL & " ORDER BY gra_no"

	'Response.Write strSQL & "<br>"

	rs.Open strSQL, conn

	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount	

    strDisplayList = ""

	if not DB_RecSetIsEmpty(rs) Then

		'strDisplayList = strDisplayList & "<h2 align=""center"">Total: " & intRecordCount & " items in this pallet.</h2>"

		For intRecord = 1 To rs.PageSize
			if (DateDiff("d",rs("date_modified"), strTodayDate) = 0) OR (DateDiff("d",rs("date_created"), strTodayDate) = 0) then
				if iRecordCount Mod 2 = 0 then
					strDisplayList = strDisplayList & "<tr class=""updated_today"">"
				else
					strDisplayList = strDisplayList & "<tr class=""updated_today_2"">"
				end if
			else
				if iRecordCount Mod 2 = 0 then
					strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
				else
					strDisplayList = strDisplayList & "<tr class=""innerdoc_2"">"
				end if
			end if

			if left(rs("gra_no"),1) = "0" then
				strDisplayList = strDisplayList & "<td><a href=""update_gra_report.asp?id=" & rs("report_id") & """ target=""_blank""><img src=""images/icon_view.png"" border=""0""></a></td>"
			else
				strDisplayList = strDisplayList & "<td><a href=""update_gra-report.asp?id=" & rs("report_id") & """ target=""_blank""><img src=""images/icon_view.png"" border=""0""></a></td>"
			end if
			strDisplayList = strDisplayList & "<td align=""center"" nowrap>" & trim(rs("created_by")) & " - " & WeekDayName(WeekDay(rs("date_created"))) & ", " & FormatDateTime(rs("date_created"),1) & "</td>"

			if left(rs("gra_no"),1) = "0" then
				strDisplayList = strDisplayList & "<td><a href=""view_gra.asp?ref=report&id=" & rs("gra_no") & """ target=""_blank"">" & trim(rs("gra_no")) & "</a></td>"
			else
				strDisplayList = strDisplayList & "<td>" & trim(rs("gra_no")) & "</td>"
			end if

			strDisplayList = strDisplayList & "<td>" & trim(rs("line_no")) & "</td>"
			strDisplayList = strDisplayList & "<td>" & trim(rs("item")) & "</td>"
			strDisplayList = strDisplayList & "<td>" & trim(rs("LIC")) & "</td>"
			strDisplayList = strDisplayList & "<td>" & trim(rs("serial_no")) & "</td>"
			strDisplayList = strDisplayList & "<td>" & trim(rs("dealer_code")) & "</td>"
			strDisplayList = strDisplayList & "<td>" & trim(rs("repair_report")) & " "
			if DateDiff("d",rs("date_created"), strTodayDate) = 0 then
				strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=""0"">"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>" & trim(rs("labour")) & "</td>"
			strDisplayList = strDisplayList & "<td>" & trim(rs("parts")) & "</td>"
			strDisplayList = strDisplayList & "<td>" & trim(rs("gst")) & "</td>"
			strDisplayList = strDisplayList & "<td>" & trim(rs("total")) & "</td>"
			strDisplayList = strDisplayList & "<td>" & trim(rs("date_received")) & "</td>"
			strDisplayList = strDisplayList & "<td>" & trim(rs("date_repaired")) & "</td>"
			strDisplayList = strDisplayList & "<td>" & trim(rs("destination")) & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"" nowrap>"
			Select Case rs("status")
				case 1
					strDisplayList = strDisplayList & "Open"
				case 2
					strDisplayList = strDisplayList & "Waiting for parts"
				case 3
					strDisplayList = strDisplayList & "To be invoiced"
				case 4
					strDisplayList = strDisplayList & "Received"
				case else
			 		strDisplayList = strDisplayList & "<font color=""green"">Completed / Exported</font>"
			end select
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>"
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "</tr>"

			intTotalLabourCount = intTotalLabourCount + rs("labour")
			intTotalPartsCount 	= intTotalPartsCount + rs("parts")
			intTotalCount 		= intTotalCount + rs("total")

			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next
	else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""18"" align=""center"">No items found in this pallet.</td></tr>"
	end if

	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"	
	strDisplayList = strDisplayList & "<td colspan=""18"" class=""recordspaging"">"
	strDisplayList = strDisplayList & "<h2>Total: " & intRecordCount & " items in this pallet</h2>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseDataBase()
end sub

sub main
	call UTL_validateLogin

	Dim strPalletNo
	strPalletNo = Trim(Request("pallet_no"))

	call getPalletDetails(strPalletNo)

    if trim(session("gra_report_initial_page"))  = "" then
    	session("gra_report_initial_page") = 1
	end if

    call displayGraReports
end sub

call main

dim strMessageText

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList

dim strTotal
%>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><img src="images/icon_pallet.jpg" border="0" alt="Pallet" /></td>
          <td valign="top"><img src="images/icon_excel.jpg" border="0" /> <a href="export_pallet.asp?pallet_no=<%= request("pallet_no") %>">Export</a>
          <h2>Pallet no: <%= Request("pallet_no") %></h2>
            <h3><%= Session("pallet_department") %><br />
              <%= Session("pallet_info") %></h3></td>
          <td valign="top"><table cellpadding="4" cellspacing="0" class="created_table">
              <tr>
                <td class="created_column_1"><strong>Created:</strong></td>
                <td class="created_column_2"><%= session("pallet_created_by") %></td>
                <td class="created_column_3"><%= displayDateFormatted(session("pallet_date_created")) %></td>
              </tr>
              <tr>
                <td><strong>Last modified:</strong></td>
                <td><%= session("pallet_modified_by") %></td>
                <td><%= displayDateFormatted(session("pallet_date_modified")) %></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <p><a href="list_gra.asp">Goods Return BASE</a> &nbsp;-&nbsp; <a href="list_gra_report.asp">Report Summaries</a> &nbsp;-&nbsp; <a href="list_gra_report_writeoffs.asp">Write Offs Report</a> &nbsp;-&nbsp; <a href="list_gra_report_exported.asp">Exported Report</a> &nbsp;-&nbsp; <span class="current_header">Pallets</span></p>
      <p><img src="images/backward_arrow.gif" border="0" /> <a href="list_pallet.asp">Back to pallet list</a></p>
      <p><font color="green"><%= strMessageText %></font></p></td>
  </tr>
  <tr>
    <td><table cellspacing="0" cellpadding="4" class="database_records">
        <tr class="innerdoctitle" align="center">
          <td></td>
          <td>Created</td>
          <td>GRA #</td>
          <td>Line</td>
          <td>Item</td>
          <td>LIC</td>
          <td>Serial no</td>
          <td>Dealer code</td>
          <td>Repair report</td>
          <td>Labour</td>
          <td>Parts</td>
          <td>GST</td>
          <td>Total</td>
          <td>Received</td>
          <td>Repaired</td>
          <td>Destination</td>
          <td>Report status</td>
          <td></td>
        </tr>
        <%= strDisplayList %>
      </table></td>
  </tr>
</table>
</body>
</html>