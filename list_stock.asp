<!--#include file="include/connection_base.asp " -->
<%
session.lcid = 2057
sub setSearch	
	select case trim(request("type"))
		case "reset"
			session("stock_search") 		= ""
			session("stock_department") 	= "2"
			session("stock_warehouse") 		= ""
			session("stock_item_type") 		= ""
			session("stock_avail") 			= ""
			session("stock_sort") 			= ""
			session("stock_cinitialpage") 	= 1
		case "search"
			session("stock_search") 		= trim(request("txtSearch"))
			session("stock_department") 	= trim(request("department"))
			session("stock_warehouse") 		= trim(request("warehouse"))
			session("stock_item_type") 		= trim(request("set"))
			session("stock_avail") 			= trim(request("avail"))
			session("stock_sort") 			= trim(request("sort"))
			session("stock_cinitialpage") 	= 1
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
	
	'strSearchTxt = trim(Request("txtSearch"))
	
    call OpenBaseDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 100
		
	strSQL = "SELECT A5SOSC, A5JZSU, A5JZKH, A5JZHS, A5JZSU - (A5JZKH + A5JZHS) AS AVAILOH, "
	strSQL = strSQL & "	A5MCHS, A5MHKS, A5MHZS, A5MCHS - (A5MHKS + A5MHZS) AS AVAILIT, A5JUSG, A5SOCD, "
	strSQL = strSQL & "	Y3SYME, Y3SOMB, Y3STIK, E2IHTN + E2IHTN * E2KZRT / 100 + E2IHTN * E2SKKR / 100 AS LIC, Y1.YINWPR AS RRP, Y2.YINWPR AS S01 "
	strSQL = strSQL & " 	FROM AF5SP "
	strSQL = strSQL & "			INNER JOIN YF3MP ON A5SOSC = Y3SOSC "
	strSQL = strSQL & "			INNER JOIN EF2SP ON A5SOSC = E2SOSC "
	'strSQL = strSQL & "			INNER JOIN YFIMP ON A5SOSC = YISOSC "		
	strSQL = strSQL & "			INNER JOIN YFIMP Y1 ON A5SOSC = Y1.YISOSC "	
	strSQL = strSQL & "			INNER JOIN YFIMP Y2 ON A5SOSC = Y2.YISOSC "
	strSQL = strSQL & "			WHERE A5SOSC LIKE '%" & UCASE(trim(session("stock_search"))) & "%' "
	'strSQL = strSQL & "				AND A5SOCD LIKE '%" & UCASE(trim(session("stock_warehouse"))) & "%' "
	if trim(session("stock_warehouse")) = "" then
		strSQL = strSQL & "				AND (A5SOCD = '3S' OR A5SOCD = '3XL' OR A5SOCD = '3T')"
	else
		strSQL = strSQL & "				AND A5SOCD = '" & UCASE(trim(session("stock_warehouse"))) & "' "
	end if
	strSQL = strSQL & "				AND A5SKKI <> 'D' "
	strSQL = strSQL & "				AND Y3SKKI <> 'D' "
	strSQL = strSQL & "				AND Y3SOSC NOT LIKE '*%' "
	strSQL = strSQL & "				AND Y3SOSC NOT LIKE '#%' "
	if trim(session("stock_department")) = "407" then
		strSQL = strSQL & "				AND LEFT(Y3GREG, 3) = '" & trim(session("stock_department")) & "' "
	else
		strSQL = strSQL & "				AND LEFT(Y3GREG, 1) = '" & trim(session("stock_department")) & "' "
	end if
	
	if trim(session("stock_avail")) <> "" then
		strSQL = strSQL & "				AND ((A5JZSU - (A5JZKH + A5JZHS)) >= 1) "
	end if
	
	strSQL = strSQL & "				AND Y3STIK LIKE '%" & UCASE(trim(session("stock_item_type"))) & "%' "
	strSQL = strSQL & "				AND E2NGTY = (SELECT E2NGTY FROM EF2SP WHERE E2SOSC = A5SOSC ORDER BY E2NGTY, E2NGTM DESC Fetch First 1 Row Only) "
	strSQL = strSQL & "				AND E2NGTM = (SELECT E2NGTM FROM EF2SP WHERE E2SOSC = A5SOSC ORDER BY E2NGTY, E2NGTM DESC Fetch First 1 Row Only) "
	'strSQL = strSQL & "				AND YISKKI <> 'D' "
	strSQL = strSQL & "				AND Y1.YISKKI <> 'D' AND Y2.YISKKI <> 'D' "
	strSQL = strSQL & "				AND E2SKKI <> 'D' "
	'strSQL = strSQL & "				AND YIUSPT = 'S50' "
	strSQL = strSQL & "				AND Y1.YIUSPT = 'S50' AND Y2.YIUSPT = 'S01' "
	strSQL = strSQL & "		ORDER BY "
	
	select case session("stock_sort")		
		case "expensive"
			strSQL = strSQL & "RRP DESC"
		case "cheapest"
			strSQL = strSQL & "RRP"	
		case "available"
			strSQL = strSQL & "AVAILOH DESC"	
		case else
			strSQL = strSQL & "A5SOSC"		
	end select
	
	'response.write strSQL & "<br>"
	
	rs.Open strSQL, conn
			
	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount
	
	Select Case Request("Action")
	    case "<<"
		    intpage = 1
			session("stock_cinitialpage") = intpage
	    case "<"
		    intpage = Request("intpage") - 1
			session("stock_cinitialpage") = intpage
			
			if session("stock_cinitialpage") < 1 then session("stock_cinitialpage") = 1
	    case ">"
		    intpage = Request("intpage") + 1
			session("stock_cinitialpage") = intpage
			
			if session("stock_cinitialpage") > intPageCount then session("stock_cinitialpage") = intPageCount
	    Case ">>"
		    intpage = intPageCount
			session("stock_cinitialpage") = intpage
    end select

    strDisplayList = ""
	
	if not DB_RecSetIsEmpty(rs) Then
	
	    rs.AbsolutePage = session("stock_cinitialpage")
	
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
				strDisplayList = strDisplayList & "<tr class=""innerdoc_2"" style=""font-size:medium"">"
			end if		
			strDisplayList = strDisplayList & "<td><strong>" & rs("A5SOSC") & "</strong></td>"
			strDisplayList = strDisplayList & "<td>" & rs("Y3SYME") & " (" & Trim(rs("Y3SOMB")) & ")</td>"
			strDisplayList = strDisplayList & "<td>" & FormatNumber(rs("RRP")) & "</td>"
			strDisplayList = strDisplayList & "<td>" & FormatNumber(rs("S01")) & "</td>"
			strDisplayList = strDisplayList & "<td></td>"
			strDisplayList = strDisplayList & "<td>" & rs("A5JZSU") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("A5JZKH") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("A5JZHS") & "</td>"
			strDisplayList = strDisplayList & "<td><strong>" & rs("AVAILOH") & "</strong></td>"	
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
        strDisplayList = "<tr class=""innerdoc""><td colspan=""16"" align=""center"">No records found.</td></tr>"
	end if
	strDisplayList = strDisplayList & "<tr class=""innerdoc"" style=""font-size:medium"">"					
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
	strDisplayList = strDisplayList & "<td colspan=""16"" align=""left"">"
	strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_stock.asp"" method=""post"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & session("stock_cinitialpage") & ">"
	
	if session("stock_cinitialpage") = 1 then
   		strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
	end if
	if session("stock_cinitialpage") = intpagecount or intRecordCount = 0 then
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">>"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">>"">"
	end if
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strSearch & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboState"" value=" & strState & ">"
    strDisplayList = strDisplayList & "<h3>Page: " & session("stock_cinitialpage") & " to " & intpagecount & "</h3>"
	strDisplayList = strDisplayList & "<h2>Search results: <u>" & intRecordCount & "</u> records.</h2>"
    strDisplayList = strDisplayList & "</form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"
	
    call CloseBaseDataBase()
end sub

sub main
	call setSearch 
	
	if trim(session("stock_cinitialpage"))  = "" then
    	session("stock_cinitialpage") = 1
	end if		
    
    call displayStockAvailability
end sub

call main

dim strDisplayList
%>
<% strSection = "stockavailability" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Stock Availability</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script>
function searchStock(){    
    var strSearch 		= document.forms[0].txtSearch.value;
	var strDept  		= document.forms[0].cboDepartment.value;
	var strWarehouse  	= document.forms[0].cboWarehouse.value;
	var strItemType  	= document.forms[0].cboItemType.value;
	var intAvail 		= document.forms[0].cboAvail.value;
	var strSort 		= document.forms[0].cboSort.value; 
    document.location.href = 'list_stock.asp?type=search&txtSearch=' + strSearch + '&department=' + strDept + '&set=' + strItemType + '&warehouse=' + strWarehouse + '&avail=' + intAvail + '&sort=' + strSort;
}
    
function resetSearch(){
	document.location.href = 'list_stock.asp?type=reset';    
}  
</script>
<meta charset="utf-8">
<title>Stock Availability</title>
</head>
<body>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><h1>Stock Availability</h1>
      <table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><p><img src="images/icon_excel.jpg" border="0" /> <a href="export_stock.asp">Export</a></p></td>
          <td valign="top"><div class="alert alert-search">
              <form name="frmSearch" id="frmSearch" method="post" action="list_stock.asp?type=search" onsubmit="searchStock()">
                <input type="text" class="form-control" name="txtSearch" maxlength="15" size="20" value="<%= request("txtSearch") %>" placeholder="Search Product code" />
                <select name="cboDepartment" onchange="searchStock()" class="form-control">
                  <option <% if session("stock_department") = "2" then Response.Write " selected" end if%> value="2">AV</option>
                  <option <% if session("stock_department") = "3" then Response.Write " selected" end if%> value="3">TRAD</option>
                  <option <% if session("stock_department") = "4" then Response.Write " selected" end if%> value="4">PRO</option>
                  <option <% if session("stock_department") = "407" then Response.Write " selected" end if%> value="407">PAISTE</option>
                </select>
                <select name="cboItemType" onchange="searchStock()" class="form-control">
                  <option <% if session("stock_item_type") = "" then Response.Write " selected" end if%> value="">All Items</option>
                  <option <% if session("stock_item_type") = "3" then Response.Write " selected" end if%> value="3">Set Items</option>
                </select>
                <select name="cboWarehouse" onchange="searchStock()" class="form-control">
                  <option <% if session("stock_warehouse") = "" then Response.Write " selected" end if%> value="">3S, 3T and 3XL</option>
                  <option <% if session("stock_warehouse") = "3S" then Response.Write " selected" end if%> value="3S">3S</option>
                  <option <% if session("stock_warehouse") = "3T" then Response.Write " selected" end if%> value="3T">3T</option>
                  <option <% if session("stock_warehouse") = "3XL" then Response.Write " selected" end if%> value="3XL">3XL</option>
                </select>
                <select name="cboAvail" onchange="searchStock()" class="form-control">
                  <option <% if session("stock_avail") = "" then Response.Write " selected" end if%> value="">All Stock</option>
                  <option <% if session("stock_avail") = "1" then Response.Write " selected" end if%> value="1">Available on-hand only</option>
                </select>
                <select name="cboSort" onchange="searchStock()" class="form-control">
                  <option <% if session("stock_sort") = "product"  then Response.Write " selected" end if %> value="product">Sort by: Product code (A-Z)</option>
                  <option <% if session("stock_sort") = "expensive" then Response.Write " selected" end if %> value="expensive">Sort by: RRP (High to Low)</option>
                  <option <% if session("stock_sort") = "cheapest" then Response.Write " selected" end if %> value="cheapest">Sort by: RRP (Low to High)</option>
                  <option <% if session("stock_sort") = "available" then Response.Write " selected" end if %> value="available">Sort by: On-hand Availability (High to Low)</option>
          		</select>
                <input type="button" name="btnSearch" value="Search" onclick="searchStock()" />
                <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
              </form>
            </div></td>
        </tr>
      </table></td>
  </tr>
</table>
<table cellspacing="0" cellpadding="5" class="database_records">
  <thead>
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
  </thead>
  <tbody>
    <%= strDisplayList %>
  </tbody>
</table>
</body>
</html>