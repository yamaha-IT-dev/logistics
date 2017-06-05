<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<!--#include file="class/clsGoodsReturnReport.asp " -->
<!--#include file="class/clsPallet.asp " -->
<% strSection = "gra" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>GRA Report Summaries</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/javascript.js"></script>
<script language="JavaScript" type="text/javascript">
function searchReports(){
    var strReportSearch = document.forms[0].txtSearch.value;
	//var strReportPallet = document.forms[0].cboPallet.value;
	var strReportMonth  = document.forms[0].cboMonth.value;
	var strReportYear  	= document.forms[0].cboYear.value;
	var strReportStatus = document.forms[0].cboStatus.value;
	var strReportSort  	= document.forms[0].cboSort.value;
    document.location.href = 'list_claims.asp?type=search&txtSearch=' + strReportSearch + '&cboMonth=' + strReportMonth + '&cboYear=' + strReportYear + '&cboStatus=' + strReportStatus + '&cboSort=' + strReportSort;
}

function resetSearch(){
	document.location.href = 'list_claims.asp?type=reset';
}

function confirmClear(theForm) {
	var reason = "";
	var blnSubmit = true;
	
	if (blnSubmit == true){
		//alert("clearing");
        theForm.command.value = 'clear';
  		//theForm.submit();

		return true;
    }
}
</script>
</head>
<body>
<%
sub setSearch
	select case Trim(Request("type"))
		case "reset"
			session("claim_search") 		= ""
			session("claim_month") 			= ""
			session("claim_year") 			= ""
			session("claim_status") 		= ""
			session("claim_sort") 			= ""
			session("claim_initial_page") 	= 1
		case "search"
			session("claim_search") 		= Trim(Request("txtSearch"))
			session("claim_month") 			= Trim(Request("cboMonth"))
			session("claim_year") 			= Trim(Request("cboYear"))
			session("claim_status") 		= Trim(Request("cboStatus"))
			session("claim_sort") 			= Trim(Request("cboSort"))
			session("claim_initial_page") 	= 1
	end select
end sub

sub displayClaims
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
	
	if session("claim_status") = "" then
		session("claim_status") = "1"
	end if
	
	if session("claim_year") = "" then
		session("claim_year") = "2014"
	end if

	if session("claim_sort") = "" then
		session("claim_sort") = "gra_no"
	end if

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation 	= 3	'adUseClient
    rs.CursorType 		= 3	'adOpenStatic
	rs.PageSize 		= 100			
			
	strSQL = "SELECT * FROM yma_gra_report "
	strSQL = strSQL & " WHERE YEAR(date_created) = '" & trim(session("claim_year")) & "' "
	if session("claim_month") <> "" then
		strSQL = strSQL & " AND MONTH(date_created) = '" & trim(session("claim_month")) & "' "
	end if
	strSQL = strSQL & " 	AND (gra_no LIKE '%" & session("claim_search") & "%' "
	strSQL = strSQL & "			OR item LIKE '%" & session("claim_search") & "%' "
	strSQL = strSQL & "			OR serial_no LIKE '%" & session("claim_search") & "%') "
	strSQL = strSQL & " 	AND status LIKE '%" & session("claim_status") & "%' "
	strSQL = strSQL & " ORDER BY " & session("claim_sort")	
	
	'Response.Write strSQL

	rs.Open strSQL, conn

	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount
	
	Select Case Request("Action")
	    case "<<"
		    intpage = 1
			session("claim_initial_page") = intpage
	    case "<"
		    intpage = Request("intpage") - 1
			session("claim_initial_page") = intpage

			if session("claim_initial_page") < 1 then session("claim_initial_page") = 1
	    case ">"
		    intpage = Request("intpage") + 1
			session("claim_initial_page") = intpage

			if session("claim_initial_page") > intPageCount then session("claim_initial_page") = IntPageCount
	    Case ">>"
		    intpage = intPageCount
			session("claim_initial_page") = intpage
    end select

    strDisplayList = ""

	if not DB_RecSetIsEmpty(rs) Then

		For intRecord = 1 To rs.PageSize
		
			strDisplayList = strDisplayList & "<form method=""post"" action=""list_claims.asp"">"
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
				strDisplayList = strDisplayList & "<td align=""center""><a href=""update_gra_report.asp?id=" & rs("report_id") & """>Edit Report</a></td>"
			else
				strDisplayList = strDisplayList & "<td align=""center""><a href=""update_gra-report.asp?id=" & rs("report_id") & """>Edit Report</a></td>"
			end if	
			strDisplayList = strDisplayList & "<td align=""center"">" & trim(rs("created_by")) & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & WeekDayName(WeekDay(rs("date_created"))) & ", " & FormatDateTime(rs("date_created"),1) & "</td>"	
			
			if left(rs("gra_no"),1) = "0" then
				strDisplayList = strDisplayList & "<td align=""center""><a href=""view_gra.asp?ref=report&id=" & rs("gra_no") & """>" & trim(rs("gra_no")) & "</a></td>"
			else
				strDisplayList = strDisplayList & "<td align=""center"">" & trim(rs("gra_no")) & "</td>"
			end if
			
			strDisplayList = strDisplayList & "<td align=""center"">" & trim(rs("line_no")) & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & trim(rs("item")) & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & trim(rs("serial_no")) & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & trim(rs("dealer_code")) & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & trim(rs("repair_report")) & " "
			if DateDiff("d",rs("date_created"), strTodayDate) = 0 then
				strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=""0"">"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & trim(rs("labour")) & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & trim(rs("parts")) & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & trim(rs("gst")) & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & trim(rs("total")) & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & trim(rs("date_received")) & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & trim(rs("date_repaired")) & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & trim(rs("destination")) & "</td>"
			'strDisplayList = strDisplayList & "<td align=""center"">" & trim(rs("pallet_no")) & "</td>"			
			strDisplayList = strDisplayList & "<td align=""center"">"
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
			strDisplayList = strDisplayList & "<td align=""center"">"
			
			if trim(rs("status")) = 3 then
				strDisplayList = strDisplayList & "<input type=""submit"" value=""Export"" />"			
			end if
			
			strDisplayList = strDisplayList & "</td>"
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
        strDisplayList = "<tr class=""innerdoc""><td colspan=""18"" align=""center"">There are no reports.</td></tr>"
	end if

	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"	
	strDisplayList = strDisplayList & "<td colspan=""18"" class=""recordspaging"">"
	strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_claims.asp"" method=""post"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & session("claim_initial_page") & ">"

	if session("claim_initial_page") = 1 then
   		strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
	end if
	if session("claim_initial_page") = intpagecount or intRecordCount = 0 then
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
    strDisplayList = strDisplayList & "Page: " & session("claim_initial_page") & " to " & intpagecount
    strDisplayList = strDisplayList & "<br />"
	strDisplayList = strDisplayList & "Search results: " & intRecordCount & " reports."
    strDisplayList = strDisplayList & "</form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseDataBase()
end sub

sub main
	call UTL_validateLogin
	call setSearch
	
    if trim(session("claim_initial_page"))  = "" then
    	session("claim_initial_page") = 1
	end if

    call displayClaims
	call listReportTotal		
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
          <td valign="top">
          <div class="alert alert-success"><img src="images/add_icon.png" border="0" align="bottom" /> <a href="add_gra-report.asp">Add Report Manually</a></div>
          <img src="images/icon_excel.jpg" border="0" /> <a href="export_gra_report.asp?search=<%= session("claim_search") %>&year=<%= session("claim_year") %>&month=<%= session("claim_month") %>&status=<%= session("claim_status") %>&sort=<%= session("claim_sort") %>">Export</a></td>
          <td valign="top"><div class="alert alert-search">
              <form name="frmSearch" id="frmSearch" action="list_claims.asp?type=search" method="post" onsubmit="searchReports()">
                <h3>Report Filter:</h3>
                GRA no / Item /Serial no:
                <input type="text" name="txtSearch" size="25" value="<%= request("txtSearch") %>" maxlength="20" />                
                <select name="cboMonth" onchange="searchReports()">
                  <option <% if session("claim_month") = "" then Response.Write " selected" end if%> value="">All months</option>
                  <option <% if session("claim_month") = "1" then Response.Write " selected" end if%> value="1">January</option>
                  <option <% if session("claim_month") = "2" then Response.Write " selected" end if%> value="2">February</option>
                  <option <% if session("claim_month") = "3" then Response.Write " selected" end if%> value="3">March</option>
                  <option <% if session("claim_month") = "4" then Response.Write " selected" end if%> value="4">April</option>
                  <option <% if session("claim_month") = "5" then Response.Write " selected" end if%> value="5">May</option>
                  <option <% if session("claim_month") = "6" then Response.Write " selected" end if%> value="6">June</option>
                  <option <% if session("claim_month") = "7" then Response.Write " selected" end if%> value="7">July</option>
                  <option <% if session("claim_month") = "8" then Response.Write " selected" end if%> value="8">August</option>
                  <option <% if session("claim_month") = "9" then Response.Write " selected" end if%> value="9">September</option>
                  <option <% if session("claim_month") = "10" then Response.Write " selected" end if%> value="10">October</option>
                  <option <% if session("claim_month") = "11" then Response.Write " selected" end if%> value="11">November</option>
                  <option <% if session("claim_month") = "12" then Response.Write " selected" end if%> value="12">December</option>
                </select>
                <select name="cboYear" onchange="searchReports()">
                  <option <% if session("claim_year") = "2014" then Response.Write " selected" end if%> value="2014">2014</option>
                  <option <% if session("claim_year") = "2013" then Response.Write " selected" end if%> value="2013">2013</option>
                  <option <% if session("claim_year") = "2012" then Response.Write " selected" end if%> value="2012">2012</option>
                  <option <% if session("claim_year") = "2011" then Response.Write " selected" end if%> value="2011">2011</option>
                  <option <% if session("claim_year") = "2010" then Response.Write " selected" end if%> value="2010">2010</option>
                  <option <% if session("claim_year") = "2009" then Response.Write " selected" end if%> value="2009">2009</option>
                  <option <% if session("claim_year") = "2008" then Response.Write " selected" end if%> value="2008">2008</option>
                </select>
                <select name="cboStatus" onchange="searchReports()">
                  <option <% if session("claim_status") = "1" then Response.Write " selected" end if%> value="1">Status: Open</option>
                  <option <% if session("claim_status") = "4" then Response.Write " selected" end if%> value="4">Status: Received</option>
                  <option <% if session("claim_status") = "2" then Response.Write " selected" end if%> value="2">Status: Waiting for parts</option>
                  <option <% if session("claim_status") = "3" then Response.Write " selected" end if%> value="3">Status: To be invoiced</option>                  
                  <option <% if session("claim_status") = "0" then Response.Write " selected" end if%> value="0" style="color:green">Status: Completed / Exported</option>
                </select>
                <select name="cboSort" onchange="searchReports()">
                  <option <% if session("claim_sort") = "gra_no" then Response.Write " selected" end if%> value="gra_no">Sort by: GRA no</option>
                  <option <% if session("claim_sort") = "date_created DESC" then Response.Write " selected" end if%> value="date_created DESC">Sort by: Date created</option>
                  <option <% if session("claim_sort") = "item" then Response.Write " selected" end if%> value="item">Sort by: Item</option>
                  <option <% if session("claim_sort") = "date_received" then Response.Write " selected" end if%> value="date_received">Sort by: Date received</option>                  
                </select>
                <input type="button" name="btnSearch" value="Search" onclick="searchReports()" />
                <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
              </form>
            </div></td>
        </tr>
      </table>
      <p><a href="list_gra.asp">Goods Return BASE</a> &nbsp;-&nbsp; <span class="current_header">Report Summaries</span> &nbsp;-&nbsp; <a href="list_gra_report_writeoffs.asp">Write Offs Report</a> &nbsp;-&nbsp; <a href="list_gra_report_exported.asp">Exported Report</a> &nbsp;-&nbsp; <a href="list_pallet.asp">Pallets</a></p>
      <form action="" name="form_clear" id="form_clear" method="post" onsubmit="return confirmClear(this)">
      <p align="right">
      	<input type="hidden" name="command" value="clear">
        <input type="submit" value="Clear" <% if session("UsrUserName") <> "jeffj" and session("UsrUserName") <> "harsonos" and session("UsrUserName") <> "matthewm" then Response.Write "disabled" end if%> />
      </p>
      </form>
      <p><font color="green"><%= strMessageText %></font></p>
      <table cellspacing="0" cellpadding="4" class="database_records">
        <tr class="innerdoctitle" align="center">
          <td></td>
          <td>Created by</td>
          <td>Date created</td>
          <td>GRA no</td>
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
          <td>Report status</td>
          <td></td>
        </tr>
        <%= strDisplayList %>
      </table>
      <%= strTotal %>      
      </td>
  </tr>
</table>
</body>
</html>