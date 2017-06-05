<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<% strSection = "gra" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Goods Return</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/javascript.js"></script>
<script language="JavaScript" type="text/javascript">
function searchItem(){    
    var strSearch 		= document.forms[0].txtSearch.value;
	var strOperator		= document.forms[0].cboOperator.value;
	var strMonth  		= document.forms[0].cboMonth.value;
	var strYear  		= document.forms[0].cboYear.value;
	var strWarehouse	= document.forms[0].cboWarehouse.value;	
	var strStatus 		= document.forms[0].cboStatus.value;
	
    document.location.href = 'list_gra_item.asp?type=search&txtSearch=' + strSearch + '&operator=' + strOperator + '&month=' + strMonth + '&year=' + strYear + '&warehouse=' + strWarehouse + '&status=' + strStatus;
}
    
function resetSearch(){
	document.location.href = 'list_gra_item.asp?type=reset';    
}  
</script>
</head>
<body>
<%
sub setSearch	
	select case Trim(Request("type"))
		case "reset"
			session("gra_item_search") 			= ""
			session("gra_item_search_operator") 	= ""
			session("gra_item_search_month") 	= ""
			session("gra_item_search_year") 		= ""
			session("gra_item_search_warehouse")	= ""			
			session("gra_item_search_status") 	= ""
			session("gra_initial_page") 	= 1
		case "search"
			session("gra_item_search") 			= Trim(Request("txtSearch"))
			session("gra_item_search_operator") 	= Trim(Request("operator"))
			session("gra_item_search_month") 	= Trim(Request("month"))
			session("gra_item_search_year") 		= Trim(Request("year"))
			session("gra_item_search_warehouse")	= Trim(Request("warehouse"))			
			session("gra_item_search_status") 	= Trim(Request("status"))
			session("gra_initial_page") 	= 1
	end select
end sub

sub displayGraItem
	dim iRecordCount
	iRecordCount = 0    
    dim strSQL
	dim intRecordCount
	dim strTodayDate	
	
	strTodayDate = FormatDateTime(Date())
	
    call OpenDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 100
	
	if trim(session("gra_item_search_year"))  = "" then
    	session("gra_item_search_year") = 2014
	end if
		
	strSQL = "SELECT DISTINCT BTUSNO, BTKASC, BTOPEC, BTHYNO, BTURKC + BTHSRC AS dealer_code, BTAHYY, BTAHYM, BTAHYD, " 
	strSQL = strSQL & "  cast(BTAHYY as varchar(4)) + right('00'+cast(BTAHYM as varchar(2)),2) + right('00'+cast(BTAHYD as varchar(2)),2) return_date, "
	strSQL = strSQL & "  BTRTST, BTUSGR, BTATT1, BTSCNO, BTGICM, BTNICM, "
	strSQL = strSQL & "  YMOPNM, Y1JSO1, Y1JSO3, Y1UBNB, Y1KOM1, Y1KNCI, Y1TELN, gra_connote, "
	strSQL = strSQL & "  	CASE "
	strSQL = strSQL & "			WHEN deduction_gra_no IS NULL THEN 'no' "
	strSQL = strSQL & "			ELSE 'yes'"
	strSQL = strSQL & "  	END AS 'deduction' "
	strSQL = strSQL & "  FROM OPENQUERY(AS400, 'SELECT BTUSNO, BTKASC, BTOPEC, BTHYNO, BTURKC, BTHSRC, BTAHYY, BTAHYM, BTAHYD, BTAHYY, BTAHYM, BTAHYD, BTRTST, BTUSGR, BTATT1, BTSCNO, BTGICM, BTNICM, BTSKKI FROM BFTEP') "
	strSQL = strSQL & "  	INNER JOIN OPENQUERY(AS400, 'SELECT BUHYNO, BUSOSC, BUCLMN, BUSIBN FROM BFUEP') ON BTHYNO = BUHYNO "
	strSQL = strSQL & " 		LEFT JOIN OPENQUERY(AS400, 'SELECT Y1KOM1, Y1JSO1, Y1JSO3, Y1KNCI, Y1UBNB, Y1TELN, Y1KOKC FROM YF1MP WHERE Y1SKKI <> ''D''') ON BTURKC + BTHSRC = Y1KOKC "
	strSQL = strSQL & " 		LEFT JOIN OPENQUERY(AS400, 'SELECT DISTINCT YMOPEC, YMOPNM, YMSKKI, YMPMID FROM YFMMP') ON BTOPEC = YMOPEC "
	strSQL = strSQL & " 		LEFT JOIN tbl_deductions on BTHYNO = deduction_gra_no "
	strSQL = strSQL & " 		LEFT JOIN yma_gra_status on BTHYNO = gra_no "
	strSQL = strSQL & " WHERE (BTHYNO LIKE '%" & session("gra_item_search") & "%' "
	strSQL = strSQL & " 		OR BTURKC LIKE '%" & UCASE(session("gra_item_search")) & "%' "
	strSQL = strSQL & " 		OR BTATT1 LIKE '%" & UCASE(session("gra_item_search")) & "%' "
	strSQL = strSQL & " 		OR YMOPNM LIKE '%" & UCASE(session("gra_item_search")) & "%' "
	strSQL = strSQL & " 		OR Y1KOM1 LIKE '%" & UCASE(session("gra_item_search")) & "%' "
	strSQL = strSQL & "			OR BUSIBN LIKE '%" & UCASE(session("gra_item_search")) & "%' "
	strSQL = strSQL & "			OR BUCLMN LIKE '%" & UCASE(session("gra_item_search")) & "%' "
	strSQL = strSQL & "			OR BUSOSC LIKE '%" & UCASE(session("gra_item_search")) & "%')"
	strSQL = strSQL & " 	AND BTSKKI <> 'D' "
	if trim(session("gra_item_search_month")) <> "" then
		strSQL = strSQL & " 	AND BTAHYM = '" & session("gra_item_search_month") & "' "
	end if
	strSQL = strSQL & " 	AND BTAHYY = '" & session("gra_item_search_year") & "' "
	strSQL = strSQL & " 	AND YMOPNM LIKE '%" & UCASE(session("gra_item_search_operator")) & "%' "
	strSQL = strSQL & " 	AND BTSCNO LIKE '%" & session("gra_item_search_warehouse") & "%' "
	strSQL = strSQL & "		AND BTRTST LIKE '%" & session("gra_item_search_status") & "%' "
	strSQL = strSQL & "		AND (YMPMID like 'LOG%' OR YMPMID like 'INT%' OR YMPMID like 'SERV%' OR YMPMID like 'CREDIT%')"
	strSQL = strSQL & "		AND (YMSKKI <> 'D' OR (YMSKKI = 'D' AND YMOPEC = 'AC')) "
	strSQL = strSQL & " 	ORDER BY BTHYNO DESC "	
	
	'Response.Write strSQL & "<br>"
	
	rs.Open strSQL, conn
			
	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount
	
	Select Case Request("Action")
	    case "<<"
		    intpage = 1
			session("gra_initial_page") = intpage
	    case "<"
		    intpage = Request("intpage") - 1
			session("gra_initial_page") = intpage
			
			if session("gra_initial_page") < 1 then session("gra_initial_page") = 1
	    case ">"
		    intpage = Request("intpage") + 1
			session("gra_initial_page") = intpage
			
			if session("gra_initial_page") > intPageCount then session("gra_initial_page") = IntPageCount
	    Case ">>"
		    intpage = intPageCount
			session("gra_initial_page") = intpage
    end select

    strDisplayList = ""
	
	if not DB_RecSetIsEmpty(rs) Then
	
	    rs.AbsolutePage = session("gra_initial_page")  
		
		For intRecord = 1 To rs.PageSize 
			if iRecordCount Mod 2 = 0 then
				strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
			else
				strDisplayList = strDisplayList & "<tr class=""innerdoc_2"">"
			end if
			strDisplayList = strDisplayList & "<td><a href=""view_gra.asp?ref=gra&id=" & rs("BTHYNO") & """><img src=""images/icon_view.png"" border=""0""></a></td>"
			strDisplayList = strDisplayList & "<td>" & trim(rs("YMOPNM")) & "</td>"
			strDisplayList = strDisplayList & "<td>"
			if rs("deduction") = "yes" then
				strDisplayList = strDisplayList & "<font style=""background-color:yellow"">" & trim(rs("BTHYNO")) & "</font>"
			else
				strDisplayList = strDisplayList & "" & trim(rs("BTHYNO")) & ""
			end if
			strDisplayList = strDisplayList & "</td>"
			'strDisplayList = strDisplayList & "<td>" & trim(rs("BUCLMN")) & "</td>"
			strDisplayList = strDisplayList & "<td>" & trim(rs("dealer_code")) & "</td>"
			strDisplayList = strDisplayList & "<td>"
			strDisplayList = strDisplayList & "<span title="" " & trim(rs("Y1JSO1")) 
			strDisplayList = strDisplayList & ", " & trim(rs("Y1JSO3")) 
			strDisplayList = strDisplayList & " " & trim(rs("Y1UBNB")) & " "">"
			strDisplayList = strDisplayList & trim(rs("Y1KOM1")) & "</span>"
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>"
			Select Case trim(rs("Y1KNCI"))
				case "01"
					strDisplayList = strDisplayList & "ACT"
				case "02"
					strDisplayList = strDisplayList & "NSW"
				case "03"
					strDisplayList = strDisplayList & "VIC"
				case "04"
					strDisplayList = strDisplayList & "QLD"
				case "05"
					strDisplayList = strDisplayList & "SA"
				case "06"
					strDisplayList = strDisplayList & "WA"
				case "07"
					strDisplayList = strDisplayList & "TAS"
				case "08"
					strDisplayList = strDisplayList & "NT"				
				case else
					strDisplayList = strDisplayList & rs("Y1KNCI")
			end select
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>" & trim(rs("BTGICM")) & " " & trim(rs("BTNICM")) & "</td>"			
			strDisplayList = strDisplayList & "<td>" & trim(rs("Y1TELN")) & "</td>"			
			strDisplayList = strDisplayList & "<td>"
			strDisplayList = strDisplayList & "" & trim(rs("BTAHYD")) & " "
			Select Case trim(rs("BTAHYM"))
				case "1"
					strDisplayList = strDisplayList & "Jan"
				case "2"
					strDisplayList = strDisplayList & "Feb"
				case "3"
					strDisplayList = strDisplayList & "Mar"
				case "4"
					strDisplayList = strDisplayList & "Apr"
				case "5"
					strDisplayList = strDisplayList & "May"
				case "6"
					strDisplayList = strDisplayList & "Jun"
				case "7"
					strDisplayList = strDisplayList & "Jul"
				case "8"
					strDisplayList = strDisplayList & "Aug"
				case "9"
					strDisplayList = strDisplayList & "Sep"
				case "10"
					strDisplayList = strDisplayList & "Oct"
				case "11"
					strDisplayList = strDisplayList & "Nov"
				case "12"
					strDisplayList = strDisplayList & "Dec"			
			end select
			strDisplayList = strDisplayList & " " & trim(rs("BTAHYY"))			
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>"
			Select Case trim(rs("BTRTST"))
				case "0"
					strDisplayList = strDisplayList & "<font color=blue>Not received</font>"
				case "1"
					strDisplayList = strDisplayList & "<font color=red>Received not credited</font>"
				case "2"
					strDisplayList = strDisplayList & "<font color=green>Credited</font>"
				case else
			 		strDisplayList = strDisplayList & trim(rs("BTRTST"))
			end select
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>"
			Select Case trim(rs("BTUSGR"))
				case "J"
					strDisplayList = strDisplayList & "<img src=""images/cope.gif"" border=""0"">"
				case "C"
					strDisplayList = strDisplayList & "Custom pickup"				
				case else
			 		strDisplayList = strDisplayList & trim(rs("BTUSGR"))
			end select
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>" & trim(rs("gra_connote")) & "</td>"
			strDisplayList = strDisplayList & "<td>" & trim(rs("BTSCNO")) & "</td>"
			strDisplayList = strDisplayList & "</tr>"
			
			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For 
		next
	else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""13"" align=""center"">There are no records.</td></tr>"
	end if
	
	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""13"" class=""recordspaging"">"
	strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_gra_item.asp"" method=""post"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & session("gra_initial_page") & ">"
	
	if session("gra_initial_page") = 1 then
   		strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
	else 
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
	end if	
	if session("gra_initial_page") = intpagecount or intRecordCount = 0 then
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">>"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">>"">"
	end if
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strSearch & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboDepartment"" value=" & strItemDepartment & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboStatus"" value=" & strStatus & ">"
    strDisplayList = strDisplayList & "<br />"
    strDisplayList = strDisplayList & "Page: " & session("gra_initial_page") & " to " & intpagecount
	strDisplayList = strDisplayList & "<br />"
	strDisplayList = strDisplayList & "<h2>Search results: " & intRecordCount & " records.</h2>"
    strDisplayList = strDisplayList & "</form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"
	
    call CloseDataBase()
end sub

sub main
	call UTL_validateLogin
	call setSearch

    if trim(session("gra_initial_page"))  = "" then
    	session("gra_initial_page") = 1
	end if

    call displayGraItem
end sub

call main

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList
%>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><img src="images/icon_gra.jpg" border="0" alt="GRA" /></td>
          <!--<td valign="top"><img src="images/icon_excel.jpg" border="0" /> <a href="#">Export</a></td>-->
          <td valign="top"><div class="alert alert-search">
              <form name="frmSearch" id="frmSearch" action="list_gra_item.asp?type=search" method="post" onsubmit="searchItem()">
                <h3>GRA Search Parameters:</h3>
                Operator / Item / Serial no / GRA no / Dealer code / Dealer / Contact / Claim no:
                <input type="text" name="txtSearch" size="25" value="<%= request("txtSearch") %>" maxlength="20" />
                <select name="cboOperator" onchange="searchItem()">
                  <option <% if session("gra_item_search_operator") = "" then Response.Write " selected" end if%> value="">All Operators</option>
                  <option <% if session("gra_item_search_operator") = "aaron" then Response.Write " selected" end if%> value="aaron">Aaron Chen</option>
                  <option <% if session("gra_item_search_operator") = "boyd" then Response.Write " selected" end if%> value="boyd">Boyd Gill</option>
                  <option <% if session("gra_item_search_operator") = "lecky" then Response.Write " selected" end if%> value="lecky">David Lecky</option>
                  <option <% if session("gra_item_search_operator") = "cooper" then Response.Write " selected" end if%> value="cooper">Daniel Cooper</option>
                  <option <% if session("gra_item_search_operator") = "whitehead" then Response.Write " selected" end if%> value="whitehead">Joshua Whitehead</option>
                  <option <% if session("gra_item_search_operator") = "scholes" then Response.Write " selected" end if%> value="scholes">Johanna Scholes</option>
                  <option <% if session("gra_item_search_operator") = "walker" then Response.Write " selected" end if%> value="walker">Lyn Walker</option>
                  <option <% if session("gra_item_search_operator") = "lapthorne" then Response.Write " selected" end if%> value="lapthorne">Mark Lapthorne</option>
                  <option <% if session("gra_item_search_operator") = "collyer" then Response.Write " selected" end if%> value="collyer">Matt Collyer</option>
                  <option <% if session("gra_item_search_operator") = "livingston" then Response.Write " selected" end if%> value="livingston">Matt Livingstone</option>
                  <option <% if session("gra_item_search_operator") = "mathew" then Response.Write " selected" end if%> value="mathew">Mathew Taylor</option>
                  <option <% if session("gra_item_search_operator") = "madden" then Response.Write " selected" end if%> value="madden">Matthew Madden</option>
                  <option <% if session("gra_item_search_operator") = "lush" then Response.Write " selected" end if%> value="lush">Stefan Lush</option>
                  <option <% if session("gra_item_search_operator") = "steers" then Response.Write " selected" end if%> value="steers">Timothy Steers</option>
                </select>
                <select name="cboMonth" onchange="searchItem()">
                  <option <% if session("gra_item_search_month") = "" then Response.Write " selected" end if%> value="">All Months (Plan Return Date)</option>
                  <option <% if session("gra_item_search_month") = "1" then Response.Write " selected" end if%> value="1">January</option>
                  <option <% if session("gra_item_search_month") = "2" then Response.Write " selected" end if%> value="2">February</option>
                  <option <% if session("gra_item_search_month") = "3" then Response.Write " selected" end if%> value="3">March</option>
                  <option <% if session("gra_item_search_month") = "4" then Response.Write " selected" end if%> value="4">April</option>
                  <option <% if session("gra_item_search_month") = "5" then Response.Write " selected" end if%> value="5">May</option>
                  <option <% if session("gra_item_search_month") = "6" then Response.Write " selected" end if%> value="6">June</option>
                  <option <% if session("gra_item_search_month") = "7" then Response.Write " selected" end if%> value="7">July</option>
                  <option <% if session("gra_item_search_month") = "8" then Response.Write " selected" end if%> value="8">August</option>
                  <option <% if session("gra_item_search_month") = "9" then Response.Write " selected" end if%> value="9">September</option>
                  <option <% if session("gra_item_search_month") = "10" then Response.Write " selected" end if%> value="10">October</option>
                  <option <% if session("gra_item_search_month") = "11" then Response.Write " selected" end if%> value="11">November</option>
                  <option <% if session("gra_item_search_month") = "12" then Response.Write " selected" end if%> value="12">December</option>
                </select>
                <select name="cboYear" onchange="searchItem()">
				<option <% if session("gra_item_search_year") = "2016" then Response.Write " selected" end if%> value="2016">2016 (Plan Return Date)</option>
				<option <% if session("gra_item_search_year") = "2015" then Response.Write " selected" end if%> value="2015">2015</option>
                  <option <% if session("gra_item_search_year") = "2014" then Response.Write " selected" end if%> value="2014">2014</option>
                  <option <% if session("gra_item_search_year") = "2013" then Response.Write " selected" end if%> value="2013">2013</option>
                  <option <% if session("gra_item_search_year") = "2012" then Response.Write " selected" end if%> value="2012">2012</option>
                  <option <% if session("gra_item_search_year") = "2011" then Response.Write " selected" end if%> value="2011">2011</option>
                  <option <% if session("gra_item_search_year") = "2010" then Response.Write " selected" end if%> value="2010">2010</option>
                  <option <% if session("gra_item_search_year") = "2009" then Response.Write " selected" end if%> value="2009">2009</option>
                  <option <% if session("gra_item_search_year") = "2008" then Response.Write " selected" end if%> value="2008">2008</option>
                </select>
                <select name="cboWarehouse" onchange="searchItem()">
                  <option <% if session("gra_item_search_warehouse") = "" then Response.Write " selected" end if%> value="">All Warehouses</option>
                  <option <% if session("gra_item_search_warehouse") = "3K" then Response.Write " selected" end if%> value="3K">3K</option>
                  <option <% if session("gra_item_search_warehouse") = "3L" then Response.Write " selected" end if%> value="3L">3L</option>
                  <option <% if session("gra_item_search_warehouse") = "3ND" then Response.Write " selected" end if%> value="3ND">3ND</option>
                  <option <% if session("gra_item_search_warehouse") = "3S" then Response.Write " selected" end if%> value="3S">3S</option>
                  <option <% if session("gra_item_search_warehouse") = "3T" then Response.Write " selected" end if%> value="3T">3T</option>
                  <option <% if session("gra_item_search_warehouse") = "3TH" then Response.Write " selected" end if%> value="3TH">3TH</option>
                  <option <% if session("gra_item_search_warehouse") = "3XL" then Response.Write " selected" end if%> value="3XL">3XL</option>
                  <option <% if session("gra_item_search_warehouse") = "4T" then Response.Write " selected" end if%> value="4T">4T</option>
                  <option <% if session("gra_item_search_warehouse") = "6T" then Response.Write " selected" end if%> value="6T">6T</option>
                </select>
                <select name="cboStatus" onchange="searchItem()">
                  <option <% if session("gra_item_search_status") = "" then Response.Write " selected" end if%> value="">All Status</option>
                  <option <% if session("gra_item_search_status") = "0" then Response.Write " selected" end if%> value="0">0 - Not received</option>
                  <option <% if session("gra_item_search_status") = "1" then Response.Write " selected" end if%> value="1">1 - Received but not credited</option>
                  <option <% if session("gra_item_search_status") = "2" then Response.Write " selected" end if%> value="2">2 - Credited</option>
                </select>
                <input type="button" name="btnSearch" value="Search" onclick="searchItem()" />
                <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
              </form>
            </div></td>
        </tr>
      </table>
      <p><span class="current_header">Goods Return BASE</span> &nbsp;-&nbsp; <a href="list_gra_report.asp">Report Summaries</a> &nbsp;-&nbsp; <a href="list_gra_report_writeoffs.asp">Write Offs Report</a> &nbsp;-&nbsp; <a href="list_gra_report_exported.asp">Exported Report</a> &nbsp;-&nbsp; <a href="list_pallet.asp">Pallets</a></p></td>
  </tr>
  <tr>
    <td><table cellspacing="0" cellpadding="5" class="database_records">
        <tr class="innerdoctitle">
          <td>&nbsp;</td>
          <td>Line</td>
          <td>Item</td>
          <td>Expected return qty</td>
          <td>Qty received</td>
          <td>Serial</td>
          <td>Return code</td>
          <td>Condition</td>
          <td>Ori invoice</td>
          <td>Credit note</td>
          <td>Amount</td>
          <td>Claim no</td>
          <td>Received</td>
        </tr>
        <%= strDisplayList %>
      </table></td>
  </tr>
</table>
</body>
</html>