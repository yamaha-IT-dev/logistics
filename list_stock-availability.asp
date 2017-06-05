<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<% strSection = "stockavailability" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Stock Availability</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/javascript.js"></script>
<script type="text/javascript" src="include/jquery.js"></script>
<script type="text/javascript" src="include/main.js"></script>
<script language="JavaScript" type="text/javascript">
function searchStock(){    
    var strSearch 		= document.forms[0].txtSearch.value;
	var strDept  		= document.forms[0].cboDepartment.value;
	var strWarehouse  	= document.forms[0].cboWarehouse.value;
	var strItemType  	= document.forms[0].cboItemType.value;
    document.location.href = 'list_stock-availability.asp?type=search&txtSearch=' + strSearch + '&cboDepartment=' + strDept + '&type=' + strItemType + '&cboWarehouse=' + strWarehouse;
}

function resetSearch(){
	document.location.href = 'list_stock-availability.asp?type=reset';
}
</script>
</head>
<body>
<%
sub setSearch	
	select case Trim(Request("type"))
		case "reset"
			session("stockavail_search") 		= ""
			session("stockavail_department") 	= "2"
			session("stockavail_warehouse") 	= ""
			session("stockavail_item_type") 	= ""
			session("stockavail_cinitialpage") 	= 1
		case "search"
			session("stockavail_search") 		= Trim(Request("txtSearch"))
			session("stockavail_department") 	= Trim(Request("cboDepartment"))
			session("stockavail_warehouse") 	= Trim(Request("cboWarehouse"))
			session("stockavail_item_type") 	= Trim(Request("cboItemType"))
			session("stockavail_cinitialpage") 	= 1
	end select
end sub

sub displayStockAvailability
	dim iRecordCount
	iRecordCount = 0
    dim strSortBy
	dim strSortItem
    'dim strSearchTxt
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
	
	dim strPageResultNumber
	dim strRecordPerPage
	dim intRecordCount
	
	dim strTodayDate
	strTodayDate = FormatDateTime(Date())
	
	'strSearchTxt = Trim(Request("txtSearch"))
	
    call OpenBaseDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 100
		
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
	
	'response.write strSQL & "<br>"
	
	rs.Open strSQL, conn
			
	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount
	
	Select Case Request("Action")
	    case "<<"
		    intpage = 1
			session("stockavail_cinitialpage") = intpage
	    case "<"
		    intpage = Request("intpage") - 1
			session("stockavail_cinitialpage") = intpage
			
			if session("stockavail_cinitialpage") < 1 then session("stockavail_cinitialpage") = 1
	    case ">"
		    intpage = Request("intpage") + 1
			session("stockavail_cinitialpage") = intpage
			
			if session("stockavail_cinitialpage") > intPageCount then session("stockavail_cinitialpage") = intPageCount
	    Case ">>"
		    intpage = intPageCount
			session("stockavail_cinitialpage") = intpage
    end select

    strDisplayList = ""
	
	if not DB_RecSetIsEmpty(rs) Then
	
	    rs.AbsolutePage = session("stockavail_cinitialpage")
	
		For intRecord = 1 To rs.PageSize 
			intOnHandsTotal 			= intOnHandsTotal + Cint(rs("A5JZSU"))
			intOnHandsTotalReserved 	= intOnHandsTotalReserved + Cint(rs("A5JZKH"))			
			intOnHandsTotalAllocated 	= intOnHandsTotalAllocated + Cint(rs("A5JZHS"))
			intOnHandsTotalAvailable 	= intOnHandsTotalAvailable + Cint(rs("AVAILOH"))
			intInTransitTotal 			= intInTransitTotal + Cint(rs("A5MCHS"))
			intInTransitTotalReserved 	= intInTransitTotalReserved + Cint(rs("A5MHKS"))
			intInTransitTotalAllocated 	= intInTransitTotalAllocated + Cint(rs("A5MHZS"))
			intInTransitTotalAvailable 	= intInTransitTotalAvailable + Cint(rs("AVAILIT"))
			intTotalBackorder 			= intTotalBackorder + Cint(rs("A5JUSG"))
			
			if iRecordCount Mod 2 = 0 then
				strDisplayList = strDisplayList & "<tr class=""innerdoc"" style=""font-size:medium"">"
			else
				strDisplayList = strDisplayList & "<tr class=""innerdoc_2""  style=""font-size:medium"">"
			end if		
			strDisplayList = strDisplayList & "<td>" & rs("A5SOSC") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("Y3SOMB") & "</td>"
			strDisplayList = strDisplayList & "<td>" & FormatNumber(rs("RRP")) & "</td>"
			strDisplayList = strDisplayList & "<td>" & FormatNumber(rs("S01")) & "</td>"
			'strDisplayList = strDisplayList & "<td>" & FormatNumber(rs("LIC")) & "</td>"
			strDisplayList = strDisplayList & "<td></td>"
			strDisplayList = strDisplayList & "<td>" & rs("A5JZSU") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("A5JZKH") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("A5JZHS") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("AVAILOH") & "</td>"	
			strDisplayList = strDisplayList & "<td></td>"		
			strDisplayList = strDisplayList & "<td>" & rs("A5MCHS") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("A5MHKS") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("A5MHZS") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("AVAILIT") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("A5JUSG") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("A5SOCD") & "</td>"		
			strDisplayList = strDisplayList & "</tr>"
			
			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For 
		next

	else
        strDisplayList = "<tr><td colspan=""16"" align=""center"" bgcolor=""white"" style=""font-size:large"">No records found.</td></tr>"
	end if
	strDisplayList = strDisplayList & "<tr bgcolor=""grey"" style=""color:white; font-size:large"">"					
	strDisplayList = strDisplayList & "<td><b>Total of this page: </b></td>"
	strDisplayList = strDisplayList & "<td></td>"
	strDisplayList = strDisplayList & "<td></td>"
	strDisplayList = strDisplayList & "<td></td>"
	strDisplayList = strDisplayList & "<td></td>"
	strDisplayList = strDisplayList & "<td><b>" & intOnHandsTotal & "</b></td>"
	strDisplayList = strDisplayList & "<td><b>" & intOnHandsTotalReserved & "</b></td>"
	strDisplayList = strDisplayList & "<td><b>" & intOnHandsTotalAllocated & "</b></td>"
	strDisplayList = strDisplayList & "<td><b>" & intOnHandsTotalAvailable & "</b></td>"
	strDisplayList = strDisplayList & "<td></td>"
	strDisplayList = strDisplayList & "<td><b>" & intInTransitTotal & "</b></td>"
	strDisplayList = strDisplayList & "<td><b>" & intInTransitTotalReserved & "</b></td>"
	strDisplayList = strDisplayList & "<td><b>" & intInTransitTotalAllocated & "</b></td>"
	strDisplayList = strDisplayList & "<td><b>" & intInTransitTotalAvailable & "</b></td>"
	strDisplayList = strDisplayList & "<td><b>" & intTotalBackorder & "</b></td>"
	strDisplayList = strDisplayList & "<td></td>"		
	strDisplayList = strDisplayList & "</tr>"			
	strDisplayList = strDisplayList & "<tr>"
	strDisplayList = strDisplayList & "<td colspan=""16"" align=""center"" bgcolor=""white"">"
	strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_stock-availability.asp"" method=""post"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & session("stockavail_cinitialpage") & ">"
	
	if session("stockavail_cinitialpage") = 1 then
   		strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
	end if
	if session("stockavail_cinitialpage") = intpagecount or intRecordCount = 0 then
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">>"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">>"">"
	end if
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strSearch & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboState"" value=" & strState & ">"
    strDisplayList = strDisplayList & "<h3>Page: " & session("stockavail_cinitialpage") & " to " & intpagecount & "</h3>"
	strDisplayList = strDisplayList & "<h2>Search results: <u>" & intRecordCount & "</u> records.</h2>"
    strDisplayList = strDisplayList & "</form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"
	
    call CloseBaseDataBase()
end sub

sub main
	call UTL_validateLogin
	call setSearch 
	
	if trim(session("stockavail_cinitialpage"))  = "" then
    	session("stockavail_cinitialpage") = 1
	end if		
    
    call displayStockAvailability
end sub

call main

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList
%>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content">
    <h1>Stock Availability</h1>
    <%= session("stockavail_department") %>
    <table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top">
            <p><img src="images/icon_excel.jpg" border="0" /> <a href="export_stock-availability.asp">Export</a></p>
            </td>
          <td valign="top"><div class="alert alert-search">
              <form name="frmSearch" id="frmSearch" method="post" action="list_stock-availability.asp?type=search" onsubmit="searchStock()">Search Product code:
                <input type="text" class="form-control" name="txtSearch" maxlength="15" size="20" value="<%= request("txtSearch") %>" placeholder="Search Product code" />
                <select id="cboDepartment" name="cboDepartment" onchange="searchStock()" class="form-control">
                  <option <% if session("stockavail_department") = "2" then Response.Write " selected" end if%> value="2">AV</option>
                  <option <% if session("stockavail_department") = "3" then Response.Write " selected" end if%> value="3">TRAD</option>
                  <option <% if session("stockavail_department") = "4" then Response.Write " selected" end if%> value="4">PRO</option>
                  <option <% if session("stockavail_department") = "407" then Response.Write " selected" end if%> value="407">PAISTE</option>
                </select>
                <select id="cboItemType" name="cboItemType" onchange="searchStock()" class="form-control">
                  <option <% if session("stockavail_item_type") = "" then Response.Write " selected" end if%> value="">All Items</option>
                  <option <% if session("stockavail_item_type") = "1" then Response.Write " selected" end if%> value="1">Set Items</option>
                </select>
                <select id="cboWarehouse" name="cboWarehouse" onchange="searchStock()" class="form-control">
                  <option <% if session("stockavail_warehouse") = "" then Response.Write " selected" end if%> value="">3S, 3T and 3XL</option>
                  <option <% if session("stockavail_warehouse") = "3S" then Response.Write " selected" end if%> value="3S">3S</option>
                  <option <% if session("stockavail_warehouse") = "3T" then Response.Write " selected" end if%> value="3T">3T</option>
                  <option <% if session("stockavail_warehouse") = "3XL" then Response.Write " selected" end if%> value="3XL">3XL</option>
                </select>
                <input type="button" name="btnSearch" value="Search" onclick="searchStock()" />
                <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
              </form>
            </div></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td><table cellspacing="0" cellpadding="5" class="database_records">
        <tr class="innerdoctitle" style="font-size:medium">
          <td><strong>Product Code</strong></td>
          <td><strong>Description</strong></td>
          <td><strong>RRP</strong></td>
          <td><strong>S01</strong></td>
          <td class="column_header" nowrap><strong><u>ON-HAND:</u></strong></td>
          <td class="column_header"><strong>Total</strong></td>
          <td class="column_header"><strong>Reserved</strong></td>
          <td class="column_header"><strong>Allocated</strong></td>
          <td class="column_header"><strong>Available</strong></td>
          <td class="column_header_alt" nowrap><strong><u>IN-TRANSIT:</u></strong></td>
          <td class="column_header_alt"><strong>Total</strong></td>
          <td class="column_header_alt"><strong>Reserved</strong></td>
          <td class="column_header_alt"><strong>Allocated</strong></td>
          <td class="column_header_alt"><strong>Available</strong></td>
          <td><strong>Backorder</strong></td>
          <td><strong>Warehouse</strong></td>
        </tr>
        <%= strDisplayList %>
      </table></td>
  </tr>
</table>
</body>
</html>