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
	var strReportDate 	= document.forms[0].cboDate.value;
	var strReportMonth  = document.forms[0].cboMonth.value;
	var strReportYear  	= document.forms[0].cboYear.value;
	var strReportStatus = document.forms[0].cboStatus.value;
	var strReportSort  	= document.forms[0].cboSort.value;
    document.location.href = 'list_gra_report.asp?type=search&txtSearch=' + strReportSearch + '&date=' + strReportDate + '&month=' + strReportMonth + '&year=' + strReportYear + '&status=' + strReportStatus + '&sort=' + strReportSort;
}

function resetSearch(){
	document.location.href = 'list_gra_report.asp?type=reset';
}

function confirmClear(theForm) {
	var reason = "";
	var blnSubmit = true;
	
	if (blnSubmit == true){

        theForm.command.value = 'clear';

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
			session("gra_report_search") 		= ""
			session("gra_report_date") 			= ""
			session("gra_report_month") 		= ""
			session("gra_report_year") 			= ""
			session("gra_report_status") 		= ""
			session("gra_report_sort") 			= ""
			session("gra_report_initial_page") 	= 1
		case "search"
			session("gra_report_search") 		= Trim(Request("txtSearch"))
			session("gra_report_date") 			= Trim(Request("date"))
			session("gra_report_month") 		= Trim(Request("month"))
			session("gra_report_year") 			= Trim(Request("year"))
			session("gra_report_status") 		= Trim(Request("status"))
			session("gra_report_sort") 			= Trim(Request("sort"))
			session("gra_report_initial_page") 	= 1
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
	
	if session("gra_report_date") = "" then
		session("gra_report_date") = "date_created"
	end if

	if session("gra_report_sort") = "" then
		session("gra_report_sort") = "gra"
	end if

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation 	= 3	'adUseClient
    rs.CursorType 		= 3	'adOpenStatic
	rs.PageSize 		= 300			
	
	strSQL = "SELECT yma_gra_report.*, BTAHYD, BTAHYM, BTAHYY FROM yma_gra_report "
	strSQL = strSQL & " LEFT JOIN OPENQUERY(AS400, 'SELECT BTHYNO, BTAHYD, BTAHYM, BTAHYY FROM BFTEP') ON BTHYNO = gra_no "
	strSQL = strSQL & " WHERE "
	strSQL = strSQL & " 	(gra_no LIKE '%" & session("gra_report_search") & "%' "
	strSQL = strSQL & "			OR dealer_code LIKE '%" & session("gra_report_search") & "%' "
	strSQL = strSQL & "			OR item LIKE '%" & session("gra_report_search") & "%' "
	strSQL = strSQL & "			OR serial_no LIKE '%" & session("gra_report_search") & "%') "
	
	if session("gra_report_month") <> "" then
		strSQL = strSQL & " AND MONTH(" & session("gra_report_date") & ") = '" & trim(session("gra_report_month")) & "' "
	end if
	
	if session("gra_report_year") <> "" then
		strSQL = strSQL & " AND YEAR(" & session("gra_report_date") & ") = '" & trim(session("gra_report_year")) & "' "
	end if
		
	strSQL = strSQL & " 	AND status LIKE '%" & session("gra_report_status") & "%' "
	strSQL = strSQL & " ORDER BY " 
	select case session("gra_report_sort")
		case "gra"
			strSQL = strSQL & "gra_no"
		case "item"
			strSQL = strSQL & "item"
		case "created_latest"
			strSQL = strSQL & "date_created DESC"
		case "created_oldest"
			strSQL = strSQL & "date_created"	
		case "received_latest"
			strSQL = strSQL & "date_received DESC"
		case "received_oldest"
			strSQL = strSQL & "date_received"
		case "repaired_latest"
			strSQL = strSQL & "date_repaired DESC"
		case "repaired_oldest"
			strSQL = strSQL & "date_repaired"	
		case "exported_latest"
			strSQL = strSQL & "date_exported DESC"
		case "exported_oldest"
			strSQL = strSQL & "date_exported"	
		case else
			strSQL = strSQL & "gra_no"
	end select	
	
	'Response.Write strSQL & "<br>"

	rs.Open strSQL, conn

	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount
	
	Select Case Request("Action")
	    case "<<"
		    intpage = 1
			session("gra_report_initial_page") = intpage
	    case "<"
		    intpage = Request("intpage") - 1
			session("gra_report_initial_page") = intpage

			if session("gra_report_initial_page") < 1 then session("gra_report_initial_page") = 1
	    case ">"
		    intpage = Request("intpage") + 1
			session("gra_report_initial_page") = intpage

			if session("gra_report_initial_page") > intPageCount then session("gra_report_initial_page") = IntPageCount
	    Case ">>"
		    intpage = intPageCount
			session("gra_report_initial_page") = intpage
    end select

    strDisplayList = ""

	if not DB_RecSetIsEmpty(rs) Then
		
		strDisplayList = strDisplayList & "<h2 align=""center"">Search results: " & intRecordCount & " reports.</h2>"
		
		For intRecord = 1 To rs.PageSize
		
			strDisplayList = strDisplayList & "<form method=""post"" action=""list_gra_report.asp"">"
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
			'strDisplayList = strDisplayList & "<td nowrap>" & trim(rs("created_by")) & " - " & WeekDayName(WeekDay(rs("date_created"))) & ", " & FormatDateTime(rs("date_created"),1) & "</td>"	
			strDisplayList = strDisplayList & "<td nowrap>" & trim(rs("created_by")) & " - " & FormatDateTime(rs("date_created"),1) & "</td>"	
			
			if left(rs("gra_no"),1) = "0" then
				strDisplayList = strDisplayList & "<td><a href=""view_gra.asp?ref=report&id=" & rs("gra_no") & """>" & trim(rs("gra_no")) & "</a></td>"
			else
				strDisplayList = strDisplayList & "<td>" & trim(rs("gra_no")) & "</td>"
			end if
			strDisplayList = strDisplayList & "<td>" & rs("BTAHYD") & "/" & rs("BTAHYM") & "/" & rs("BTAHYY") & "</td>"
			strDisplayList = strDisplayList & "<td>" & trim(rs("line_no")) & "</td>"
			strDisplayList = strDisplayList & "<td><strong>" & trim(rs("item")) & "</strong></td>"
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
			strDisplayList = strDisplayList & "<td><strong>" & trim(rs("date_repaired")) & "</strong></td>"
			strDisplayList = strDisplayList & "<td>" & trim(rs("destination")) & " " & trim(rs("pallet_no")) & "</td>"
			'strDisplayList = strDisplayList & "<td>" & trim(rs("pallet_no")) & "</td>"			
			strDisplayList = strDisplayList & "<td nowrap>"
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
			 		strDisplayList = strDisplayList & "<font color=""green"">Completed / Exported</font> - " & rs("date_exported")
			end select
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>"
			
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
	strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_gra_report.asp"" method=""post"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & session("gra_report_initial_page") & ">"

	if session("gra_report_initial_page") = 1 then
   		strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
	end if
	if session("gra_report_initial_page") = intpagecount or intRecordCount = 0 then
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
    strDisplayList = strDisplayList & "Page: " & session("gra_report_initial_page") & " to " & intpagecount	
	strDisplayList = strDisplayList & "<br />"
	strDisplayList = strDisplayList & "<h2>Search results: " & intRecordCount & " reports.</h2>"
    strDisplayList = strDisplayList & "</form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseDataBase()
end sub

'----------------------------------------------------------------------------------------
' SET GRA REPORT INVOICE_EXPORTED FLAG TO 1
'----------------------------------------------------------------------------------------
sub updateGraReportExportedFlag
	dim strSQL
	dim intID
	intID 	= Request.Form("report_id")
	
	Call OpenDataBase()

	strSQL = "UPDATE yma_gra_report SET "
	strSQL = strSQL & "invoice_exported = '1',"
	strSQL = strSQL & "date_exported = getdate(),"
	strSQL = strSQL & "status = '0' WHERE report_id = " & intID

	'response.Write strSQL
	
	on error resume next
	conn.Execute strSQL

	if err <> 0 then
		strMessageText = err.description
	else
		strMessageText = "The GRA report has been succesfully exported."
	end if

	Call CloseDataBase()
end sub

'----------------------------------------------------------------------------------------
' INSERT GRA REPORT TO TEMP TABLE IN BASE
'----------------------------------------------------------------------------------------

sub exportGraReportToBase
	dim strSQL
	
	dim strGraNo
	dim intLineNo
	dim strItem
	dim strSerialNo
	dim strDealerCode
	dim strRepairReport
	dim strComments
	dim strWarrantyCode
	dim intLabour
	dim intParts
	dim intGST
	
	strGraNo 		= Request.Form("gra_no")
	intLineNo 		= Request.Form("line_no")
	strItem 		= Request.Form("item")
	strSerialNo 	= Request.Form("serial_no")
	strDealerCode 	= Request.Form("dealer_code")
	strRepairReport = Request.Form("repair_report")
	strComments 	= Request.Form("comments")
	strWarrantyCode	= Request.Form("warranty_code")
	intLabour 		= Request.Form("labour")
	intParts 		= Request.Form("parts")
	intGST 			= Request.Form("gst")
	
	call OpenDataBase()
	
	'TEST
	'strSQL = "INSERT INTO openquery(s1027cfg, 'SELECT * FROM OFPAP_TEST') ("
	
	'LIVE
	strSQL = "INSERT INTO openquery(s1027cfg, 'SELECT * FROM OFPAP') ("
	
	strSQL = strSQL & " OPCLIM, " 'gra no
	strSQL = strSQL & " OPCCFL, " 'cheque / credit note flag = 0
	strSQL = strSQL & " OPRCTI, " 'recepient created tax invoice = 1
	strSQL = strSQL & " OPSISC, " 'vendor code = EXCEL
	strSQL = strSQL & " OPURKC, " 'Dealer code = first 6
	strSQL = strSQL & " OPJURC, " 'Dealer code = bill to
	strSQL = strSQL & " OPHSRC, " 'Dealer code = ship to
	strSQL = strSQL & " OPSOSC, " 'item
	strSQL = strSQL & " OPOMDF, " 'old model flag = 1
	strSQL = strSQL & " OPSIBN, " 'serial no
	strSQL = strSQL & " OPSSE, " 'purchase date
	strSQL = strSQL & " OPRTLN, " 'dealer code
	strSQL = strSQL & " OPCOMP, " 'fault text
	strSQL = strSQL & " OPTECR, " 'repair report
	strSQL = strSQL & " OPEXCC, " 'external comment code
	strSQL = strSQL & " OPEXCM, " 'external comment text
	strSQL = strSQL & " OPINCC, " 'internal comment code
	strSQL = strSQL & " OPINCM, " 'date upload = todays date
	strSQL = strSQL & " OPWARC, " 'warranty code
	strSQL = strSQL & " OPLACH, " 'labour
	strSQL = strSQL & " OPPACH, " 'parts
	strSQL = strSQL & " OPOTCH "  'gst
	
	strSQL = strSQL & "	) VALUES ("
	
	strSQL = strSQL & "'XL" & strGraNo & "/" & intLineNo & "',"
	strSQL = strSQL & "'0',"
	strSQL = strSQL & "'1',"
	strSQL = strSQL & "'EXCEL',"
	strSQL = strSQL & "'',"
	strSQL = strSQL & "'',"
	strSQL = strSQL & "'',"
	strSQL = strSQL & "'" & strItem & "',"
	strSQL = strSQL & "'1',"
	strSQL = strSQL & "'" & strSerialNo & "',"
	strSQL = strSQL & "right('0' + convert(varchar(2), day(getdate())),2)+ right('0' + convert(varchar(2), month(getdate())),2) +  right(convert(varchar(4), year(getdate())),2),"
	strSQL = strSQL & "'" & strDealerCode & "',"
	strSQL = strSQL & "'" & strGraNo & "',"
	strSQL = strSQL & "LEFT('" & strRepairReport & "',50),"
	strSQL = strSQL & "'',"
	strSQL = strSQL & "'',"
	strSQL = strSQL & "'',"
	strSQL = strSQL & "'UPLOADED ' + convert(varchar(10), getdate(), 103),"
	strSQL = strSQL & "'" & strWarrantyCode & "',"
	strSQL = strSQL & "'" & FormatNumber(intLabour, 2, 0, 0, 0) & "',"
	strSQL = strSQL & "'" & FormatNumber(intParts, 2, 0, 0, 0) & "',"
	strSQL = strSQL & "'" & FormatNumber(intGST, 2, 0, 0, 0) & "')"
	
	'response.Write strSQL
	on error resume next
	conn.Execute strSQL
	
	if err <> 0 then
		strMessageText = err.description
	else
		call updateGraReportExportedFlag
		'strMessageText = "Successfully inserted into Temp table"
	end if 
	
	call CloseDataBase()
end sub

'----------------------------------------------------------------------------------------
' CLEAR TEMP TABLE IN BASE
'----------------------------------------------------------------------------------------
sub clearTempTable
	dim strSQL
	
    call OpenDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	
	strSQL = "DELETE openquery(s1027cfg, 'SELECT * FROM OFPAP_EXCEL')"	'
	
	'response.Write strSQL
	
	rs.Open strSQL, conn
	
	Set rs = nothing
	
	if err <> 0 then
		strMessageText = err.description
	else 
		strMessageText = "Temp table has been cleared."
	end if 
	
    call CloseDataBase()
end sub

sub main
	call UTL_validateLogin
	call setSearch
	call getAllPalletList
	
    if trim(session("gra_report_initial_page"))  = "" then
    	session("gra_report_initial_page") = 1
	end if

    call displayGraReports
	call listReportTotal
	
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then	
		Select Case Trim(Request("command"))
			case "export_gra_report"			
				call exportGraReportToBase
				call displayGraReports	
				call listReportTotal
			case "clear"
				call clearTempTable
				'response.write "CLEARING..."
		end select
	end if
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
          <td valign="top"><div class="alert alert-success"><img src="images/add_icon.png" border="0" align="bottom" /> <a href="add_gra-report.asp">Add Report Manually</a></div>
            <img src="images/icon_excel.jpg" border="0" /> <a href="export_gra_report.asp">Export</a></td>
          <td valign="top"><div class="alert alert-search">
              <form name="frmSearch" id="frmSearch" action="list_gra_report.asp?type=search" method="post" onsubmit="searchReports()">
                <h3>Report Filter:</h3>
                GRA / Item / Serial / Dealer Code:
                <input type="text" name="txtSearch" size="25" value="<%= request("txtSearch") %>" maxlength="20" />
                <select name="cboDate" onchange="searchReports()">                 
                  <option <% if session("gra_report_date") = "date_created" then Response.Write " selected" end if%> value="date_created">Date Created (Report)</option>
                  <option <% if session("gra_report_date") = "date_received" then Response.Write " selected" end if%> value="date_received">Date Received</option>
                  <option <% if session("gra_report_date") = "date_repaired" then Response.Write " selected" end if%> value="date_repaired">Date Repaired</option>
                  <option <% if session("gra_report_date") = "date_exported" then Response.Write " selected" end if%> value="date_exported">Date Exported</option>                  
                </select>
                <select name="cboMonth" onchange="searchReports()">
                  <option <% if session("gra_report_month") = "" then Response.Write " selected" end if%> value="">All months - <%= session("gra_report_date") %></option>
                  <option <% if session("gra_report_month") = "1" then Response.Write " selected" end if%> value="1">January - <%= session("gra_report_date") %></option>
                  <option <% if session("gra_report_month") = "2" then Response.Write " selected" end if%> value="2">February - <%= session("gra_report_date") %></option>
                  <option <% if session("gra_report_month") = "3" then Response.Write " selected" end if%> value="3">March - <%= session("gra_report_date") %></option>
                  <option <% if session("gra_report_month") = "4" then Response.Write " selected" end if%> value="4">April - <%= session("gra_report_date") %></option>
                  <option <% if session("gra_report_month") = "5" then Response.Write " selected" end if%> value="5">May - <%= session("gra_report_date") %></option>
                  <option <% if session("gra_report_month") = "6" then Response.Write " selected" end if%> value="6">June - <%= session("gra_report_date") %></option>
                  <option <% if session("gra_report_month") = "7" then Response.Write " selected" end if%> value="7">July - <%= session("gra_report_date") %></option>
                  <option <% if session("gra_report_month") = "8" then Response.Write " selected" end if%> value="8">August - <%= session("gra_report_date") %></option>
                  <option <% if session("gra_report_month") = "9" then Response.Write " selected" end if%> value="9">September - <%= session("gra_report_date") %></option>
                  <option <% if session("gra_report_month") = "10" then Response.Write " selected" end if%> value="10">October - <%= session("gra_report_date") %></option>
                  <option <% if session("gra_report_month") = "11" then Response.Write " selected" end if%> value="11">November - <%= session("gra_report_date") %></option>
                  <option <% if session("gra_report_month") = "12" then Response.Write " selected" end if%> value="12">December - <%= session("gra_report_date") %></option>
                </select>
                <select name="cboYear" onchange="searchReports()">
                  <option <% if session("gra_report_year") = "" then Response.Write " selected" end if%> value="">All years - <%= session("gra_report_date") %></option>
				  <option <% if session("gra_report_year") = "2017" then Response.Write " selected" end if%> value="2017">2017 - <%= session("gra_report_date") %></option>
				  <option <% if session("gra_report_year") = "2016" then Response.Write " selected" end if%> value="2016">2016 - <%= session("gra_report_date") %></option>
				  <option <% if session("gra_report_year") = "2015" then Response.Write " selected" end if%> value="2015">2015 - <%= session("gra_report_date") %></option>
                  <option <% if session("gra_report_year") = "2014" then Response.Write " selected" end if%> value="2014">2014 - <%= session("gra_report_date") %></option>
                  <option <% if session("gra_report_year") = "2013" then Response.Write " selected" end if%> value="2013">2013 - <%= session("gra_report_date") %></option>
                  <option <% if session("gra_report_year") = "2012" then Response.Write " selected" end if%> value="2012">2012 - <%= session("gra_report_date") %></option>
                  <option <% if session("gra_report_year") = "2011" then Response.Write " selected" end if%> value="2011">2011 - <%= session("gra_report_date") %></option>
                  <option <% if session("gra_report_year") = "2010" then Response.Write " selected" end if%> value="2010">2010 - <%= session("gra_report_date") %></option>
                  <option <% if session("gra_report_year") = "2009" then Response.Write " selected" end if%> value="2009">2009 - <%= session("gra_report_date") %></option>
                  <option <% if session("gra_report_year") = "2008" then Response.Write " selected" end if%> value="2008">2008 - <%= session("gra_report_date") %></option>
                </select>
                <select name="cboStatus" onchange="searchReports()">
                  <option <% if session("gra_report_status") = "" then Response.Write " selected" end if%> value="">All status</option>	
                  <option <% if session("gra_report_status") = "1" then Response.Write " selected" end if%> value="1">Status: Open</option>
                  <option <% if session("gra_report_status") = "4" then Response.Write " selected" end if%> value="4">Status: Received</option>
                  <option <% if session("gra_report_status") = "2" then Response.Write " selected" end if%> value="2">Status: Waiting for parts</option>
                  <option <% if session("gra_report_status") = "3" then Response.Write " selected" end if%> value="3">Status: To be invoiced</option>
                  <option <% if session("gra_report_status") = "0" then Response.Write " selected" end if%> value="0" style="color:green">Status: Completed / Exported</option>
                </select>
                <select name="cboSort" onchange="searchReports()">
                  <option <% if session("gra_report_sort") = "gra" then Response.Write " selected" end if%> value="gra">Sort by: GRA no</option>                  
                  <option <% if session("gra_report_sort") = "item" then Response.Write " selected" end if%> value="item">Sort by: Item</option>
                  <option <% if session("gra_report_sort") = "created_latest" then Response.Write " selected" end if%> value="created_latest">Sort by: Date Created (New - Old)</option>
                  <option <% if session("gra_report_sort") = "created_oldest" then Response.Write " selected" end if%> value="created_oldest">Sort by: Date Created (Old - New)</option>
                  <option <% if session("gra_report_sort") = "received_latest" then Response.Write " selected" end if%> value="received_latest">Sort by: Date Received (New - Old)</option>
                  <option <% if session("gra_report_sort") = "received_oldest" then Response.Write " selected" end if%> value="received_oldest">Sort by: Date Received (Old - New)</option>
                  <option <% if session("gra_report_sort") = "repaired_latest" then Response.Write " selected" end if%> value="repaired_latest">Sort by: Date Repaired (New - Old)</option>
                  <option <% if session("gra_report_sort") = "repaired_oldest" then Response.Write " selected" end if%> value="repaired_oldest">Sort by: Date Repaired (Old - New)</option>
                  <option <% if session("gra_report_sort") = "exported_latest" then Response.Write " selected" end if%> value="exported_latest">Sort by: Date Exported (New - Old)</option>
                  <option <% if session("gra_report_sort") = "exported_oldest" then Response.Write " selected" end if%> value="exported_oldest">Sort by: Date Exported (Old - New)</option>
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
      <%= strTotal %></td>
  </tr>
  <tr>
    <td><table cellspacing="0" cellpadding="8" class="database_records">
    <thead>
        <tr>
          <td></td>
          <td>Report created</td>
          <td>GRA</td>
          <td>GRA generated</td>
          <td>Line</td>
          <td>Model no</td>
          <td>Serial</td>
          <td>Dealer</td>
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
        </thead>
        <tbody>
        <%= strDisplayList %>
        </tbody>
      </table></td>
  </tr>
</table>
</body>
</html>