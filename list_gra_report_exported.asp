<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<!--#include file="class/clsGoodsReturnReport.asp " -->
<!--#include file="class/clsPallet.asp " -->
<% strSection = "gra" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>GRA Exported Report</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/javascript.js"></script>
<script language="JavaScript" type="text/javascript">
function searchReports(){
    var strReportSearch = document.forms[0].txtSearch.value;
	var strReportStart  = document.forms[0].txtStart.value;
	var strReportEnd  	= document.forms[0].txtEnd.value;
	var strReportSort  	= document.forms[0].cboSort.value;
    document.location.href = 'list_gra_report_exported.asp?type=search&txtSearch=' + strReportSearch + '&txtStart=' + strReportStart + '&txtEnd=' + strReportEnd + '&cboSort=' + strReportSort;
}

function resetSearch(){
	document.location.href = 'list_gra_report_exported.asp?type=reset';
}
</script>
</head>
<body>
<%
sub setSearch
	select case Trim(Request("type"))
		case "reset"
			session("gra_exported_report_search") 		= ""
			session("gra_exported_report_start") 		= ""
			session("gra_exported_report_end") 			= ""
			session("gra_exported_report_sort") 		= ""
			session("gra_exported_report_initial_page") = 1
		case "search"
			session("gra_exported_report_search") 		= Trim(Request("txtSearch"))
			session("gra_exported_report_start") 		= Trim(Request("txtStart"))
			session("gra_exported_report_end") 			= Trim(Request("txtEnd"))
			session("gra_exported_report_sort") 		= Trim(Request("cboSort"))
			session("gra_exported_report_initial_page") = 1
	end select
end sub

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
	
	if session("gra_exported_report_year") = "" then
		session("gra_exported_report_year") = "2014"
	end if

	if session("gra_exported_report_sort") = "" then
		session("gra_exported_report_sort") = "gra_no"
	end if

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation 	= 3	'adUseClient
    rs.CursorType 		= 3	'adOpenStatic
	rs.PageSize 		= 1000
	
	strSQL = "SELECT yma_gra_report.*, BTAHYD, BTAHYM, BTAHYY FROM yma_gra_report "
	strSQL = strSQL & " LEFT JOIN OPENQUERY(AS400, 'SELECT BTHYNO, BTAHYD, BTAHYM, BTAHYY FROM BFTEP') ON BTHYNO = gra_no "
	strSQL = strSQL & " WHERE (date_exported BETWEEN CONVERT(datetime,'" & trim(session("gra_exported_report_start")) & " 00:00:00',103)" & " "
	strSQL = strSQL & " 	AND CONVERT(datetime,'" & trim(session("gra_exported_report_end")) & " 23:59:59',103)" & ") "
	strSQL = strSQL & " 	AND (gra_no LIKE '%" & session("gra_exported_report_search") & "%' "
	strSQL = strSQL & "			OR item LIKE '%" & session("gra_exported_report_search") & "%' "
	strSQL = strSQL & "			OR serial_no LIKE '%" & session("gra_exported_report_search") & "%') "
	strSQL = strSQL & " 	AND status = '0' "
	strSQL = strSQL & " ORDER BY " & session("gra_exported_report_sort")	
	
	'Response.Write strSQL

	rs.Open strSQL, conn

	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount
	
	Select Case Request("Action")
	    case "<<"
		    intpage = 1
			session("gra_exported_report_initial_page") = intpage
	    case "<"
		    intpage = Request("intpage") - 1
			session("gra_exported_report_initial_page") = intpage

			if session("gra_exported_report_initial_page") < 1 then session("gra_exported_report_initial_page") = 1
	    case ">"
		    intpage = Request("intpage") + 1
			session("gra_exported_report_initial_page") = intpage

			if session("gra_exported_report_initial_page") > intPageCount then session("gra_exported_report_initial_page") = IntPageCount
	    Case ">>"
		    intpage = intPageCount
			session("gra_exported_report_initial_page") = intpage
    end select

    strDisplayList = ""

	if not DB_RecSetIsEmpty(rs) Then
		strDisplayList = strDisplayList & "<h3 align=""center"">Search results: " & intRecordCount & " reports.</h3>"
	
		For intRecord = 1 To rs.PageSize
		
			strDisplayList = strDisplayList & "<form method=""post"" action=""list_gra_report_exported.asp"">"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""command"" value=""export_gra_report"">"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""report_id"" value=""" & trim(rs("report_id")) & """>"				
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""gra_no"" value=""" & trim(rs("gra_no")) & """>"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""line_no"" value=""" & trim(rs("line_no")) & """>"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""item"" value=""" & trim(rs("item")) & """>"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""serial_no"" value=""" & trim(rs("serial_no")) & """>"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""dealer_code"" value=""" & trim(rs("dealer_code")) & """>"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""repair_report"" value=""" & trim(rs("repair_report")) & """>"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""comments"" value=""" & trim(rs("comments")) & """>"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""warranty_code"" value=""" & trim(rs("gra_warranty_code")) & """>"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""labour"" value=""" & trim(rs("labour")) & """>"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""parts"" value=""" & trim(rs("parts")) & """>"
			strDisplayList = strDisplayList & "<input type=""hidden"" name=""gst"" value=""" & trim(rs("gst")) & """>"
				
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
				strDisplayList = strDisplayList & "<td><a href=""update_gra_report.asp?id=" & rs("report_id") & """><img src=""images/icon_view.png"" border=""0""></a></td>"
			else
				strDisplayList = strDisplayList & "<td><a href=""update_gra-report.asp?id=" & rs("report_id") & """><img src=""images/icon_view.png"" border=""0""></a></td>"
			end if	
			'strDisplayList = strDisplayList & "<td>" & trim(rs("created_by")) & "</td>"
			'strDisplayList = strDisplayList & "<td>" & WeekDayName(WeekDay(rs("date_created"))) & ", " & FormatDateTime(rs("date_created"),1) & "</td>"	
			
			if left(rs("gra_no"),1) = "0" then
				strDisplayList = strDisplayList & "<td><a href=""view_gra.asp?ref=exported&id=" & rs("gra_no") & """>" & trim(rs("gra_no")) & "</a></td>"
			else
				strDisplayList = strDisplayList & "<td>" & trim(rs("gra_no")) & "</td>"
			end if
			strDisplayList = strDisplayList & "<td>" & rs("BTAHYD") & "/" & rs("BTAHYM") & "/" & rs("BTAHYY") & "</td>"
			strDisplayList = strDisplayList & "<td>" & trim(rs("line_no")) & "</td>"
			strDisplayList = strDisplayList & "<td>" & trim(rs("item")) & "</td>"
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
			strDisplayList = strDisplayList & "<td><strong>" & trim(rs("total")) & "</strong></td>"
			strDisplayList = strDisplayList & "<td>" & trim(rs("date_received")) & "</td>"
			strDisplayList = strDisplayList & "<td>" & trim(rs("date_repaired")) & "</td>"
			strDisplayList = strDisplayList & "<td>" & trim(rs("destination")) & "</td>"		
			
			strDisplayList = strDisplayList & "<td>" & WeekDayName(WeekDay(rs("date_exported"))) & ", " & FormatDateTime(rs("date_exported"),1) & "</td>"
			strDisplayList = strDisplayList & "</tr>"
			strDisplayList = strDisplayList & "</form>"
			
			intTotalLabourCount = intTotalLabourCount + rs("labour")
			intTotalPartsCount 	= intTotalPartsCount + rs("parts")
			intTotalCount 		= intTotalCount + rs("total")
						
			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next
	else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""17"" align=""center"">There are no reports.</td></tr>"
	end if

	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"	
	strDisplayList = strDisplayList & "<td colspan=""17"" class=""recordspaging"">"
	strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_gra_report_exported.asp"" method=""post"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & session("gra_exported_report_initial_page") & ">"

	if session("gra_exported_report_initial_page") = 1 then
   		strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
	end if
	if session("gra_exported_report_initial_page") = intpagecount or intRecordCount = 0 then
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">>"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">>"">"
	end if
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strReportSearch & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboDamageType"" value=" & strReportType & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboSort"" value=" & strReportSort & ">"
	strDisplayList = strDisplayList & "<br />"
    strDisplayList = strDisplayList & "Page: " & session("gra_exported_report_initial_page") & " to " & intpagecount
    strDisplayList = strDisplayList & "<br />"
	strDisplayList = strDisplayList & "<h2>Search results: " & intRecordCount & " reports.</h2>"
    strDisplayList = strDisplayList & "</form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseDataBase()
end sub

sub main
	call UTL_validateLogin
	call setSearch
	call getAllPalletList
	
    if trim(session("gra_exported_report_initial_page"))  = "" then
    	session("gra_exported_report_initial_page") = 1
	end if

    call displayGraReports
	call listExportedReportTotal
end sub

call main

dim strMessageText

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList
dim strDealerResultList
dim strAllPalletList

dim strTotal
%>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><img src="images/icon_gra.jpg" border="0" alt="Damage Stocks" /></td>
          <td valign="top"><img src="images/icon_excel.jpg" border="0" /> <a href="export_gra_report_exported.asp?search=<%= session("gra_exported_report_search") %>&start=<%= session("gra_exported_report_start") %>&end=<%= session("gra_exported_report_end") %>&sort=<%= session("gra_exported_report_sort") %>">Export</a>
            <p align="right"><img src="images/icon_printer.gif" border="0" /> <a href="javascript:PrintThisPage()">Printable version</a></p></td>
          <td valign="top"><div class="alert alert-search">
              <form name="frmSearch" id="frmSearch" action="list_gra_report_exported.asp?type=search" method="post" onsubmit="searchReports()">
                <h3>Exported Report Filter:</h3>
                GRA no / Item / Serial no:
                <input type="text" name="txtSearch" size="25" value="<%= request("txtSearch") %>" maxlength="20" />
                <label>Date exported start:</label>
                <input type="text" id="txtStart" name="txtStart" size="10" value="<%= request("txtStart") %>" maxlength="10" />
                <label>Date exported end:</label>
                <input type="text" id="txtEnd" name="txtEnd" size="10" value="<%= request("txtEnd") %>" maxlength="10" />
                <select name="cboSort" onchange="searchReports()">
                  <option <% if session("gra_exported_report_sort") = "gra_no" then Response.Write " selected" end if%> value="gra_no">Sort by: GRA no</option>
                  <option <% if session("gra_exported_report_sort") = "date_created DESC" then Response.Write " selected" end if%> value="date_created DESC">Sort by: Date created</option>
                  <option <% if session("gra_exported_report_sort") = "item" then Response.Write " selected" end if%> value="item">Sort by: Item</option>
                  <option <% if session("gra_exported_report_sort") = "date_received" then Response.Write " selected" end if%> value="date_received">Sort by: Date received</option>
                  <option <% if session("gra_exported_report_sort") = "date_exported" then Response.Write " selected" end if%> value="date_exported">Sort by: Date exported</option>
                </select>
                <input type="button" name="btnSearch" value="Search" onclick="searchReports()" />
                <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
              </form>
            </div></td>
        </tr>
      </table>
      <p><a href="list_gra.asp">Goods Return BASE</a> &nbsp;-&nbsp; <a href="list_gra_report.asp">Report Summaries</a> &nbsp;-&nbsp; <a href="list_gra_report_writeoffs.asp">Write Offs Report</a> &nbsp;-&nbsp; <span class="current_header">Exported Report</span> &nbsp;-&nbsp; <a href="list_pallet.asp">Pallets</a></p>
      <p><font color="green"><%= strMessageText %></font></p>
      <div id="contentstart">
        <h2>Exported between: <u><%= session("gra_exported_report_start") %></u> and <u><%= session("gra_exported_report_end") %></u></h2>
        <%= strTotal %>
        <table cellspacing="0" cellpadding="8" class="database_records">
        <thead>
          <tr>
            <td></td>
            <td>GRA no</td>
            <td>GRA generated</td>
            <td>Line</td>
            <td>Item</td>
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
            <td>Exported</td>
          </tr>
          </thead>
          <tbody>
          <%= strDisplayList %>
          </tbody>
        </table>
      </div></td>
  </tr>
</table>
<script type="text/javascript" src="include/moment.js"></script> 
<script type="text/javascript" src="include/pikaday.js"></script> 
<script type="text/javascript">	
	var picker = new Pikaday(
    {
        field: document.getElementById('txtStart'),		
        firstDay: 1,
        minDate: new Date('2000-01-01'),
        maxDate: new Date('2020-12-31'),
        yearRange: [2008,2020],
		format: 'DD/MM/YYYY'
    });
	
	var picker = new Pikaday(
    {
        field: document.getElementById('txtEnd'),		
        firstDay: 1,
        minDate: new Date('2000-01-01'),
        maxDate: new Date('2020-12-31'),
        yearRange: [2008,2020],
		format: 'DD/MM/YYYY'
    });		
	
</script>
</body>
</html>