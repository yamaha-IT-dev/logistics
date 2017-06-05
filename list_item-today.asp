<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<% strSection = "roadshow" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Roadshow Items (Updated Today)</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/javascript.js"></script>
<script language="JavaScript" type="text/javascript">
function searchItem(){
    var strSearch = document.forms[0].txtSearch.value;
	var strItemDepartment  = document.forms[0].cboDepartment.value;
	var strCategory  = document.forms[0].cboCategory.value;
	var strItemOwner  = document.forms[0].cboOwner.value;
	var strSkuType  = document.forms[0].cboSkuType.value;
	var strItemOrigin  = document.forms[0].cboOrigin.value;
	var strTransit  = document.forms[0].cboTransit.value;
	var strReturnTo  = document.forms[0].cboReturnTo.value;
    //if (strSearch != ''){
     	document.location.href = 'list_item-today.asp?type=search&txtSearch=' + strSearch + '&cboOwner=' + strItemOwner + '&cboReturnTo=' + strReturnTo + '&cboOrigin=' + strItemOrigin + '&cboDepartment=' + strItemDepartment + '&cboSkuType=' + strSkuType + '&cboTransit=' + strTransit + '&cboCategory=' + strCategory;
    //}
}

function resetSearch(){
	document.location.href = 'list_item-today.asp?type=reset';
}
</script>
</head>
<body>
<%
session.lcid = 2057

sub setSearch
	select case Trim(Request("type"))
		case "reset"
			session("strSearch") = ""
			session("strItemDepartment") = ""
			session("strItemOwner") = ""
			session("strCategory") = ""
			session("strSkuType") = ""
			session("strItemOrigin") = ""
			session("strTransit") = ""
			session("strReturnTo") = ""
			'session("cinitialPage") = 1
		case "search"
			session("strSearch") = trim(Request("txtSearch"))
			session("strItemDepartment") = request("cboDepartment")
			session("strItemOwner") = request("cboOwner")
			session("strCategory") = request("cboCategory")
			session("strSkuType") = request("cboSkuType")
			session("strItemOrigin") = request("cboOrigin")
			session("strTransit") = request("cboTransit")
			session("strReturnTo") = request("cboReturnTo")
			'session("cinitialPage") = 1
	end select
end sub

sub displayItem

    dim strSortBy
	dim strSortItem
    'Dim strSearchTxt
    dim strSQL
	'dim strDepartment
	'dim strOwner
	'dim strOrigin
	'dim strReturnTo

	dim strPageResultNumber
	dim strRecordPerPage
	dim intRecordCount

	dim strTodayDate
	strTodayDate = FormatDateTime(Date())

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 800

	strSQL = "SELECT * FROM yma_item WHERE category LIKE '%" & session("strCategory") & "%' AND department LIKE '%" & session("strItemDepartment") & "%' AND sku_type LIKE '%" & session("strSkuType") & "%' AND transit LIKE '%" & session("strTransit") & "%' AND return_to LIKE '%" & session("strReturnTo") & "%' AND origin LIKE '%" & session("strItemOrigin") & "%' AND owner LIKE '%" & session("strItemOwner") & "%' AND (product_code LIKE '%" & session("strSearch") & "%' OR description LIKE '%" & session("strSearch") & "%' OR pallet_no LIKE '%" & session("strSearch") & "%' OR loading_sequence LIKE '%" & session("strSearch") & "%') ORDER BY item_id"

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

	    'rs.AbsolutePage = session("cinitialPage")

		For intRecord = 1 To rs.PageSize
			if (DateDiff("d",rs("date_modified"), strTodayDate) = 0) or (DateDiff("d",rs("date_created"), strTodayDate) = 0) then
				strDisplayList = strDisplayList & "<tr class=""updated_today"">"
			'else
			'	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
			'end if

			'strDisplayList = strDisplayList & "<tr class=""innerdoc"" onMouseOver=""this.style.backgroundColor='#e8eef7';"" onMouseOut=""this.style.backgroundColor='#FFFFFF';"">"
			strDisplayList = strDisplayList & "<td align=""center"" nowrap><a href=""update_item.asp?id=" & rs("item_id") & """>Edit</a></td>"
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
			'if rs("status") = 1 then
			'	strDisplayList = strDisplayList & "<td align=""center"">Open</td>"
			'else
			'	strDisplayList = strDisplayList & "<td class=""green_text"">Completed</td>"
			'end if

			'strDisplayList = strDisplayList & "<td align=""center"" nowrap></td>"
			strDisplayList = strDisplayList & "<td align=""center""><a onclick=""return confirm('Are you sure you want to delete " & rs("product_code") & " ?');"" href='delete_item.asp?id=" & rs("item_id") & "'><img src=""images/btn_delete.gif"" border=""0""></a></td>"

			'strDisplayList = strDisplayList & "</tr>"
			end if
			rs.movenext

			If rs.EOF Then Exit For
		next

	else
        strDisplayList = "<tr class=innerdoc><td colspan=22 align=center>There are no active items.</td></tr>"
	end if

	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""22"" class=""recordspaging"">"
	strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_item-today.asp"" method=""post"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & session("cinitialPage") & ">"

	'if session("cinitialPage") = 1 then
   	'	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
    '	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
	'else
	'	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
    '	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
	'end if
	'if session("cinitialPage") = intpagecount or intRecordCount = 0 then
    '	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">"">"
    '	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">>"">"
	'else
	'	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">"">"
    '	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">>"">"
	'end if
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strSearch & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboDepartment"" value=" & strItemDepartment & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboOwner"" value=" & strOwner & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboCategory"" value=" & strCategory & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboSkuType"" value=" & strSkuType & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboOrigin"" value=" & strOrigin & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboTransit"" value=" & strTransit & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboReturnTo"" value=" & strReturnTo & ">"
    'strDisplayList = strDisplayList & "<input type=""hidden"" name=""order"" value=" & strSortBy & ">"
   'strDisplayList = strDisplayList & "<br />"
    'strDisplayList = strDisplayList & "Page: " & session("cinitialPage") & " to " & intpagecount
	'strDisplayList = strDisplayList & "<br />"
	'strDisplayList = strDisplayList & "Search results: " & intRecordCount & " records."
    strDisplayList = strDisplayList & "</form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseDataBase()
end sub

sub main
	call UTL_validateLogin
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
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><h2>Focus Roadshow 2011 Product List (Updated Today)</h2>

      <!--<p>Final Product List Change: <span style="color:#F00"><strong>TBA</strong></span> - <em>Please note that after this date you will not be able to update the Product List.</em></p>-->
      <img src="images/forward_arrow.gif" border="0" /> <a href="add_item.asp">Add NEW Roadshow Item</a>
       <p><a href="list_item.asp">All</a> | Updated Today</p>
      <!--<p align="right"><img src="images/icon_excel.jpg" border="0" /> <a href="export_item.asp">Export</a></p>-->
      <form name="frmSearch" id="frmSearch">
        <p>Search Item Code / Description / Pallet No / Loading Sequence:
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
          <select name="cboCategory">
          	<option value="">All Categories</option>
                <option <% if session("strCategory") = "Acoustic Drums" then Response.Write " selected" end if%> value="Acoustic Drums">Acoustic Drums</option>
                <option <% if session("strCategory") = "Brass & Woodwind" then Response.Write " selected" end if%> value="Brass & Woodwind">Brass & Woodwind</option>
                <option <% if session("strCategory") = "CA" then Response.Write " selected" end if%> value="CA">CA</option>
                <option <% if session("strCategory") = "Digital Pianos" then Response.Write " selected" end if%> value="Digital Pianos">Digital Pianos</option>
                <option <% if session("strCategory") = "Electronic Drums" then Response.Write " selected" end if%> value="Electronic Drums">Electronic Drums</option>
                <option <% if session("strCategory") = "Guitars" then Response.Write " selected" end if%> value="Guitars">Guitars</option>
                <option <% if session("strCategory") = "MPP" then Response.Write " selected" end if%> value="MPP">MPP</option>
                <option <% if session("strCategory") = "Paiste" then Response.Write " selected" end if%> value="Paiste">Paiste</option>
                <option <% if session("strCategory") = "Percussions & Strings" then Response.Write " selected" end if%> value="Percussions & Strings">Percussions & Strings</option>
                <option <% if session("strCategory") = "Pianos" then Response.Write " selected" end if%> value="Pianos">Pianos</option>
                <option <% if session("strCategory") = "POS" then Response.Write " selected" end if%> value="POS">POS</option>
                <option <% if session("strCategory") = "Portable Keyboards" then Response.Write " selected" end if%> value="Portable Keyboards">Portable Keyboards</option>
                <option <% if session("strCategory") = "Pro Audio" then Response.Write " selected" end if%> value="Pro Audio">Pro Audio</option>
                <option <% if session("strCategory") = "SYDE" then Response.Write " selected" end if%> value="SYDE">SYDE</option>
                <option <% if session("strCategory") = "VOX" then Response.Write " selected" end if%> value="VOX">VOX</option>
                <option <% if session("strCategory") = "Other" then Response.Write " selected" end if%> value="Other">Other</option>
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
          <select name="cboTransit">
            <option value="">All Freights</option>
            <option <% if session("strTransit") = "Sea Freight" then Response.Write " selected" end if%> value="Sea Freight">Sea Freight</option>
            <option <% if session("strTransit") = "Air Freight" then Response.Write " selected" end if%> value="Air Freight">Air Freight</option>
            <option <% if session("strTransit") = "NA" then Response.Write " selected" end if%> value="NA">Unknown</option>
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
          <td>&nbsp;</td>
          <td><b>ID</b></td>
          <td><b>Owner</b></td>
          <td><b>Item Code</b></td>
          <td><b>Group</b></td>
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
          <td><b>Transit</b></td>
          <td><b>Type</b></td>
          <td><b>For Sale</b></td>
          <td><b>After-show Return to</b></td>
          <td><b>Invoice #</b></td>
          <td><b>Pallet #</b></td>
          <td><b>Loading Sequence</b></td>
          <td>&nbsp;</td>
        </tr>
        <%= strDisplayList %>
      </table></td>
  </tr>
</table>
</body>
</html>