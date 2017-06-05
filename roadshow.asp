<!--#include file="include/connection_it.asp " -->
<% strSection = "roadshow" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Roadshow Items</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/javascript.js"></script>
<script language="JavaScript" type="text/javascript">
function searchItem(){    
    var strSearch = document.forms[0].txtSearch.value;
	var strItemDepartment  = document.forms[0].cboDepartment.value;
	var strItemOwner  = document.forms[0].cboOwner.value;
	var strSkuType  = document.forms[0].cboSkuType.value;
	var strItemOrigin  = document.forms[0].cboOrigin.value;
	var strReturnTo  = document.forms[0].cboReturnTo.value;
    //if (strSearch != ''){
     	document.location.href = 'roadshow.asp?type=search&txtSearch=' + strSearch + '&cboOwner=' + strItemOwner + '&cboReturnTo=' + strReturnTo + '&cboOrigin=' + strItemOrigin + '&cboDepartment=' + strItemDepartment + '&cboSkuType=' + strSkuType;
    //}            
}
    
function resetSearch(){
	document.location.href = 'roadshow.asp?type=reset';    
}  
</script>
</head>
<body>
<%
sub setSearch	
	select case Trim(Request("type"))
		case "reset"
			session("strSearch") = ""
			session("strItemDepartment") = ""
			session("strItemOwner") = ""
			session("strSkuType") = ""
			session("strItemOrigin") = ""
			session("strReturnTo") = ""
			session("cinitialPage") = 1
		case "search"
			session("strSearch") = trim(Request("txtSearch"))
			session("strItemDepartment") = request("cboDepartment")
			session("strItemOwner") = request("cboOwner")
			session("strSkuType") = request("cboSkuType")
			session("strItemOrigin") = request("cboOrigin")
			session("strReturnTo") = request("cboReturnTo")
			session("cinitialPage") = 1
	end select
end sub

sub displayItem	
	
    Dim strSortBy
	dim strSortItem
    Dim strSearchTxt
    dim strSQL
	dim strDepartment
	dim strOwner
	dim strOrigin
	dim strReturnTo
	
	dim strPageResultNumber
	dim strRecordPerPage
	dim intRecordCount
	
	dim strTodayDate
	strTodayDate = FormatDateTime(Date())
	
	'strSearchTxt = trim(Request("txtSearch"))
	'strDepartment = trim(Request("cboDepartment"))
	'strOwner = trim(Request("cboOwner"))
	'strOrigin = trim(Request("cboOrigin"))
	'strReturnTo = trim(Request("cboDelivery"))
	
    call OpenDataBase()
	
	set rs = Server.CreateObject("ADODB.recordset")
	
	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 100
	
	'strSQL = "SELECT * FROM yma_items WHERE department LIKE '%" & strDepartment & "%' AND return_to LIKE '%" & strReturnTo & "%' AND origin LIKE '%" & strOrigin & "%' AND owner LIKE '%" & strOwner & "%' AND status = 1 AND (product_code LIKE '%" & strSearchTxt & "%' OR description LIKE '%" & strSearchTxt & "%') ORDER BY item_id"
	strSQL = "SELECT * FROM yma_item WHERE department LIKE '%" & session("strItemDepartment") & "%' AND sku_type LIKE '%" & session("strSkuType") & "%' AND return_to LIKE '%" & session("strReturnTo") & "%' AND origin LIKE '%" & session("strItemOrigin") & "%' AND owner LIKE '%" & session("strItemOwner") & "%' AND (product_code LIKE '%" & session("strSearch") & "%' OR description LIKE '%" & session("strSearch") & "%') ORDER BY item_id"
			
	'Response.Write strSQL & "<br>"
	
	rs.Open strSQL, conn
			
	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount
	
	Select Case Request("Action")
	    case "<<"
		    intpage = 1
			session("cinitialPage") = intpage
	    case "<"
		    intpage = Request("intpage") - 1
			session("cinitialPage") = intpage
			
			if session("cinitialPage") < 1 then session("cinitialPage") = 1
	    case ">"
		    intpage = Request("intpage") + 1
			session("cinitialPage") = intpage
			
			if session("cinitialPage") > intPageCount then session("cinitialPage") = IntPageCount
	    Case ">>"
		    intpage = intPageCount
			session("cinitialPage") = intpage	    
    end select

    strDisplayList = ""
	
	if not DB_RecSetIsEmpty(rs) Then
	
	    rs.AbsolutePage = session("cinitialPage")  
	
		For intRecord = 1 To rs.PageSize 
			if (DateDiff("d",rs("date_modified"), strTodayDate) = 0) or (DateDiff("d",rs("date_created"), strTodayDate) = 0) then
				strDisplayList = strDisplayList & "<tr class=""updated_today"">"
			else
				strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
			end if
		
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("item_id") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("owner") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("product_code") & ""
			if DateDiff("d",rs("date_created"), strTodayDate) = 0 then
				strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=""0"">"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("item_group") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("description") & "</td>"	
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("category") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("department") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">$" & rs("rrp") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("sku_type") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("prototype") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("quantity") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("packaging") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("source") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("origin") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("available") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("transit") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("type") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("pre_sold") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("return_to") & "</td>"			
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("invoice_no") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("pallet_no") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("loading_sequence") & "</td>"
			rs.movenext
				
			If rs.EOF Then Exit For 
		next

	else
        strDisplayList = "<tr class=innerdoc><td colspan=22 align=center>There are no active items.</td></tr>"
	end if
	
	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""22"" class=""recordspaging"">"
	strDisplayList = strDisplayList & "<form name=""MovePage"" action=""roadshow.asp"" method=""post"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & session("cinitialPage") & ">"
	
	if session("cinitialPage") = 1 then
   		strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
	else 
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
	end if	
	if session("cinitialPage") = intpagecount or intRecordCount = 0 then
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">>"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">>"">"
	end if
	
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strSearchTxt & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboDepartment"" value=" & strDepartment & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboOwner"" value=" & strOwner & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboOrigin"" value=" & strOrigin & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboDelivery"" value=" & strDelivery & ">"
    strDisplayList = strDisplayList & "<br />"
    strDisplayList = strDisplayList & "Page: " & session("cinitialPage") & " to " & intpagecount
	strDisplayList = strDisplayList & "<br />"
	strDisplayList = strDisplayList & "Search results: " & intRecordCount & " records."
    strDisplayList = strDisplayList & "</form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"
	
    call CloseDataBase()
end sub

sub main
	'call UTL_validateRoadshowLogin
	call setSearch
	'Response.Write "<br>Previous: " & Request.ServerVariables("HTTP_REFERER") & Request.Querystring
    if trim(session("cinitialPage"))  = "" then
    	session("cinitialPage") = 1
	end if		
    
    call displayItem
    

end sub

call main

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList
dim strDealerResultList
dim strStateList
dim strSalesManagerList
%>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table">
  <tr>
    <td class="first_content"><h2>Focus Roadshow 2011 Product List</h2>
    <p>All | <a href="roadshow_today.asp">Updated Today</a></p>
      <p><em>Please refer any product queries to <strong>Jaclyn Williams</strong>.</em></p>
      <p align="right"><img src="images/icon_excel.jpg" border="0" /> <a href="export_item.asp">Export</a></p>
      <form name="frmSearch" id="frmSearch">
        <p>Search Item Code / Description:
          <input type="text" name="txtSearch" size="15" value="<%= request("txtSearch") %>" />
          <select name="cboDepartment">
            <option value="">All Depts</option>
            <option <% if session("strItemDepartment") = "pro" then Response.Write " selected" end if%> value="pro">Pro</option>
            <option <% if session("strItemDepartment") = "trad" then Response.Write " selected" end if%> value="trad">Trad</option>
          </select>
          <select name="cboOwner">
            <option value="">All Staff</option>
            <option <% if session("strItemOwner") = "cameront" then Response.Write " selected" end if%> value="cameront">Cameron Tait</option>
            <option <% if session("strItemOwner") = "felixe" then Response.Write " selected" end if%> value="felixe">Felix Elliot</option>
            <option <% if session("strItemOwner") = "jamesh" then Response.Write " selected" end if%> value="jamesh">James Harvey</option>
            <option <% if session("strItemOwner") = "jamieg" then Response.Write " selected" end if%> value="jamieg">Jamie Goff</option>
            <option <% if session("strItemOwner") = "nathanb" then Response.Write " selected" end if%> value="nathanb">Nathan Biggin</option>
            <option <% if session("strItemOwner") = "shaunm" then Response.Write " selected" end if%> value="shaunm">Shaun McMahon</option>
            <option <% if session("strItemOwner") = "stevenv" then Response.Write " selected" end if%> value="stevenv">Steven Vranch</option>
          </select>
          <select name="cboSkuType">
            <option value="">All SKU Types</option>
            <option <% if session("strSkuType") = "A Sku" then Response.Write " selected" end if%> value="A Sku">A Sku</option>
            <option <% if session("strSkuType") = "B Sku" then Response.Write " selected" end if%> value="B Sku">B Sku</option>
            <option <% if session("strSkuType") = "New" then Response.Write " selected" end if%> value="New">New</option>
          </select>
          <select name="cboOrigin">
            <option value="">All Origins</option>
            <option <% if session("strItemOrigin") = "Kagan" then Response.Write " selected" end if%> value="Kagan">Kagan</option>
            <option <% if session("strItemOrigin") = "Head Office" then Response.Write " selected" end if%> value="Head Office">Head Office</option>
            <option <% if session("strItemOrigin") = "3K" then Response.Write " selected" end if%> value="3K">3K</option>
            <option <% if session("strItemOrigin") = "3S" then Response.Write " selected" end if%> value="3S">3S</option>
          </select>
          <select name="cboReturnTo">
            <option value="">Aftershow Return to</option>
            <option <% if session("strReturnTo") = "Head Office" then Response.Write " selected" end if%> value="Head Office">Head Office</option>
            <option <% if session("strReturnTo") = "Dealer" then Response.Write " selected" end if%> value="Dealer">Dealer</option>
            <option <% if session("strReturnTo") = "Kagan" then Response.Write " selected" end if%> value="Kagan">Kagan</option>
            <option <% if session("strReturnTo") = "3S" then Response.Write " selected" end if%> value="3S">3S</option>
          </select>
          <input type="button" name="btnSearch" value="Search" onclick="searchItem()" />
          <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
      </form></td>
  </tr>
  <tr>
    <td class="database_column"><table cellspacing="0" cellpadding="4" class="database_records">
        <tr class="innerdoctitle" align="center">
          <td><b>ID</b></td>
          <td><b>Owner</b></td>
          <td><b>Item Code</b></td>
          <td><b>Item Group</b></td>
          <td><b>Description</b></td>
          <td><b>Category</b></td>
          <td><b>Dept</b></td>
          <td><b>RRP</b></td>
          <td><b>SKU Type</b></td>
          <td><b>Prototype</b></td>
          <td><b>Qty</b></td>
          <td><b>Packaging</b></td>
          <td><b>Source</b></td>
          <td><b>Origin</b></td>
          <td><b>Available</b></td>
          <td><b>In Transit</b></td>
          <td><b>Type</b></td>
          <td><b>For Sale</b></td>
          <td><b>After-show Return to</b></td>
          <td><b>Invoice No</b></td>
          <td><b>Pallet No</b></td>
          <td><b>Loading Sequence</b></td>
        </tr>
        <%= strDisplayList %>
      </table></td>
  </tr>
</table>
</body>
</html>