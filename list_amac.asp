<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<% strSection = "amac" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>AMAC 2014</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/javascript.js"></script>
<script language="JavaScript" type="text/javascript">
function searchItem(){
    var strSearch 			= document.forms[0].txtSearch.value;
	var strItemDepartment  	= document.forms[0].cboDepartment.value;
	var strCategory  		= document.forms[0].cboCategory.value;	
	//var strSkuType  		= document.forms[0].cboSkuType.value;
	var strItemOrigin  		= document.forms[0].cboOrigin.value;
	//var strTransit  		= document.forms[0].cboTransit.value;
	var strReturnTo  		= document.forms[0].cboReturnTo.value;
    //if (strSearch != ''){
     document.location.href = 'list_amac.asp?type=search&txtSearch=' + strSearch + '&cboReturnTo=' + strReturnTo + '&cboOrigin=' + strItemOrigin + '&cboDepartment=' + strItemDepartment + '&cboCategory=' + strCategory;
    //}
}

function resetSearch(){
	document.location.href = 'list_amac.asp?type=reset';
}
</script>
</head>
<body>
<%
sub setSearch
	select case Trim(Request("type"))
		case "reset"
			session("strSearch") 		= ""
			session("strItemDepartment") = ""
			session("strCategory") 		= ""
			session("strSkuType") 		= ""
			session("strItemOrigin") 	= ""
			'session("strTransit") 		= ""
			session("strReturnTo") 		= ""
			session("roadshow_initial_page") = 1
		case "search"
			session("strSearch") 		= trim(Request("txtSearch"))
			session("strItemDepartment") = request("cboDepartment")
			session("strCategory") 		= request("cboCategory")
			session("strSkuType") 		= request("cboSkuType")
			session("strItemOrigin") 	= request("cboOrigin")
			'session("strTransit") 		= request("cboTransit")
			session("strReturnTo") 		= request("cboReturnTo")
			session("roadshow_initial_page") = 1
	end select
end sub

sub displayRoadshow
	dim iRecordCount
	iRecordCount = 0
	
    dim strSortBy
	dim strSortItem
    'dim strSearchTxt
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

	'strSearchTxt = trim(Request("txtSearch"))
	'strDepartment = trim(Request("cboDepartment"))
	'strOwner = trim(Request("cboOwner"))
	'strOrigin = trim(Request("cboOrigin"))
	'strReturnTo = trim(Request("cboDelivery"))

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 200

	strSQL = "SELECT * FROM tbl_amac "
	strSQL = strSQL & "	WHERE category LIKE '%" & session("strCategory") & "%' "
	strSQL = strSQL & "		AND department LIKE '%" & session("strItemDepartment") & "%' "
	'strSQL = strSQL & "		AND sku_type LIKE '%" & session("strSkuType") & "%' "
	'strSQL = strSQL & "		AND transit LIKE '%" & session("strTransit") & "%' "
	strSQL = strSQL & "		AND return_to LIKE '%" & session("strReturnTo") & "%' "
	strSQL = strSQL & "		AND origin LIKE '%" & session("strItemOrigin") & "%' "
	'strSQL = strSQL & "		AND owner LIKE '%" & session("strItemOwner") & "%' "
	strSQL = strSQL & "		AND (product_code LIKE '%" & session("strSearch") & "%' "
	strSQL = strSQL & "			OR description LIKE '%" & session("strSearch") & "%') "
	'strSQL = strSQL & "			OR pallet_no LIKE '%" & session("strSearch") & "%' "
	'strSQL = strSQL & "			OR loading_sequence LIKE '%" & session("strSearch") & "%') "
	strSQL = strSQL & "	ORDER BY item_id"

	'Response.Write strSQL & "<br>"

	rs.Open strSQL, conn

	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount

	Select Case Request("Action")
	    case "<<"
		    intpage = 1
			session("roadshow_initial_page") = intpage
	    case "<"
		    intpage = Request("intpage") - 1
			session("roadshow_initial_page") = intpage

			if session("roadshow_initial_page") < 1 then session("roadshow_initial_page") = 1
	    case ">"
		    intpage = Request("intpage") + 1
			session("roadshow_initial_page") = intpage

			if session("roadshow_initial_page") > intPageCount then session("roadshow_initial_page") = intPageCount
	    Case ">>"
		    intpage = intPageCount
			session("roadshow_initial_page") = intpage
    end select

    strDisplayList = ""

	if not DB_RecSetIsEmpty(rs) Then

	    rs.AbsolutePage = session("roadshow_initial_page")

		For intRecord = 1 To rs.PageSize
			if (DateDiff("d",rs("date_modified"), strTodayDate) = 0) then
				strDisplayList = strDisplayList & "<tr class=""updated_today"">"
			else
				if iRecordCount Mod 2 = 0 then
					strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
				else
					strDisplayList = strDisplayList & "<tr class=""innerdoc_2"">"
				end if
			end if
			
			if session("UsrUserName") = "johannas" then
				strDisplayList = strDisplayList & "<td nowrap><a href=""update_amac.asp?id=" & rs("item_id") & """><img src=""images/icon_view.png"" border=""0""</a></td>"
			else
				strDisplayList = strDisplayList & "<td></td>"
			end if
			strDisplayList = strDisplayList & "<td>" & rs("item_id") & "</td>"
			
			strDisplayList = strDisplayList & "<td>" & rs("department") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("item_group") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("category") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("product_code") & ""
			if DateDiff("d",rs("date_created"), strTodayDate) = 0 then
				strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=""0"">"
			end if
			strDisplayList = strDisplayList & "</td>"
			'strDisplayList = strDisplayList & "<td>" & rs("item_group") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("description") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("rrp") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("quantity") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("sku_type") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("prototype") & "</td>"			
			strDisplayList = strDisplayList & "<td>" & rs("packaging") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("source") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("origin") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("available") & "</td>"			
			strDisplayList = strDisplayList & "<td>" & rs("transit") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("type") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("available_for_sale") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("pre_sold") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("return_to") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("displayed") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("fb_completed") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("invoice_no") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("pallet_no") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("loading_sequence") & "</td>"
			strDisplayList = strDisplayList & "<td>"
			if rs("logistics_action") = 1 then
				strDisplayList = strDisplayList & "<img src=""images/tick.gif"">"
			else
				strDisplayList = strDisplayList & "<img src=""images/cross.gif"">"
			end if
			strDisplayList = strDisplayList & "</td>"
			'strDisplayList = strDisplayList & "<td><a onclick=""return confirm('Are you sure you want to delete " & rs("product_code") & " ?');"" href='delete_roadshow.asp?id=" & rs("item_id") & "'><img src=""images/btn_delete.gif"" border=""0""></a></td>"
			strDisplayList = strDisplayList & "<td></td>"
			strDisplayList = strDisplayList & "</tr>"

			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next

	else
        strDisplayList = "<tr class=innerdoc><td colspan=""27"" align=""center"">No items found.</td></tr>"
	end if

	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""27"" class=""recordspaging"">"
	strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_amac.asp"" method=""post"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & session("roadshow_initial_page") & ">"

	if session("roadshow_initial_page") = 1 then
   		strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
	end if
	if session("roadshow_initial_page") = intpagecount or intRecordCount = 0 then
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">>"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">>"">"
	end if
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strSearch & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboDepartment"" value=" & strItemDepartment & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboOwner"" value=" & strOwner & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboCategory"" value=" & strCategory & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboSkuType"" value=" & strSkuType & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboOrigin"" value=" & strOrigin & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboTransit"" value=" & strTransit & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboReturnTo"" value=" & strReturnTo & ">"
    'strDisplayList = strDisplayList & "<input type=""hidden"" name=""order"" value=" & strSortBy & ">"
    strDisplayList = strDisplayList & "<br />"
    strDisplayList = strDisplayList & "Page: " & session("roadshow_initial_page") & " to " & intpagecount
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

    if trim(session("roadshow_initial_page"))  = "" then
    	session("roadshow_initial_page") = 1
	end if

    call displayRoadshow
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
    <td class="first_content"><h2>AMAC 2014</h2>
      <p align="right"><img src="images/icon_excel.jpg" border="0" /> <a href="export_amac.asp">Export all</a></p>         
      <div class="alert alert-search">
        <form name="frmSearch" id="frmSearch" action="list_amac.asp?type=search" method="post" onsubmit="searchItem()">
          <h3>Search Parameters:</h3>
          Item / Description / Pallet no / Loading sequence:
          <input type="text" name="txtSearch" size="15" value="<%= request("txtSearch") %>" />
          
          <select name="cboDepartment" onchange="searchItem()">
            <option value="">All Depts</option>
            <option <% if session("strItemDepartment") = "ca" then Response.Write " selected" end if%> value="ca">CA</option>
            <option <% if session("strItemDepartment") = "pro" then Response.Write " selected" end if%> value="pro">Pro</option>
            <option <% if session("strItemDepartment") = "trad" then Response.Write " selected" end if%> value="trad">Trad</option>
          </select>
          <select name="cboCategory" onchange="searchItem()">
            <option value="">All Categories</option>
            <option <% if session("strCategory") = "Acoustic Drums" then Response.Write " selected" end if%> value="Acoustic Drums">Acoustic Drums</option>
            <option <% if session("strCategory") = "Brass" then Response.Write " selected" end if%> value="Brass">Brass &amp; Woodwind</option>
            <option <% if session("strCategory") = "CA" then Response.Write " selected" end if%> value="CA">CA</option>
            <option <% if session("strCategory") = "Digital Pianos" then Response.Write " selected" end if%> value="Digital Pianos">Digital Pianos</option>
            <option <% if session("strCategory") = "Electronic Drums" then Response.Write " selected" end if%> value="Electronic Drums">Electronic Drums</option>
            <option <% if session("strCategory") = "Guitars" then Response.Write " selected" end if%> value="Guitars">Guitars</option>
            <option <% if session("strCategory") = "MPP" then Response.Write " selected" end if%> value="MPP">MPP</option>
            <option <% if session("strCategory") = "Paiste" then Response.Write " selected" end if%> value="Paiste">Paiste</option>
            <option <% if session("strCategory") = "Percussion" then Response.Write " selected" end if%> value="Percussion">Percussion &amp; Strings</option>
            <option <% if session("strCategory") = "Pianos" then Response.Write " selected" end if%> value="Pianos">Pianos</option>
            <option <% if session("strCategory") = "Portable Keyboards" then Response.Write " selected" end if%> value="Portable Keyboards">Portable Keyboards</option>
            <option <% if session("strCategory") = "Pro Audio" then Response.Write " selected" end if%> value="Pro Audio">Pro Audio</option>
            <option <% if session("strCategory") = "SYDE" then Response.Write " selected" end if%> value="SYDE">SYDE</option>
            <option <% if session("strCategory") = "VOX" then Response.Write " selected" end if%> value="VOX">VOX</option>
          </select>
          
          <select name="cboOrigin" onchange="searchItem()">
            <option value="">All Origins</option>
            <option <% if session("strItemOrigin") = "3T" then Response.Write " selected" end if%> value="3T">3T</option>
            <option <% if session("strItemOrigin") = "Head Office" then Response.Write " selected" end if%> value="Head Office">Head Office</option>
            <option <% if session("strItemOrigin") = "Vox" then Response.Write " selected" end if%> value="Vox">Direct from Vox</option>
            <option <% if session("strItemOrigin") = "YCJ" then Response.Write " selected" end if%> value="YCJ">Direct from YCJ</option>
          </select>
          
          <select name="cboReturnTo" onchange="searchItem()">
            <option value="">Aftershow Return to</option>
            <option <% if session("strReturnTo") = "3S" then Response.Write " selected" end if%> value="3S">3S</option>
            <option <% if session("strReturnTo") = "3T" then Response.Write " selected" end if%> value="3T">3T</option>
            <option <% if session("strReturnTo") = "Head Office" then Response.Write " selected" end if%> value="Head Office">Head Office</option>
            <option <% if session("strReturnTo") = "Dealer" then Response.Write " selected" end if%> value="Dealer">Dealer</option>
          </select>
          <input type="button" name="btnSearch" value="Search" onclick="searchItem()" />
          <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
        </form>
      </div></td>
  </tr>
  <tr>
    <td><table cellspacing="0" cellpadding="4" class="database_records">
        <tr class="innerdoctitle">
          <td>&nbsp;</td>
          <td>ID</td>
          <td>Dept</td>
          <td>Group</td>
          <td>Category</td>
          <td>Product code</td>
          <td>Description</td>
          <td>RRP</td>
          <td>Qty</td>
          <td>SKU</td>
          <td>Proto</td>          
          <td>Packaging</td>
          <td>Source</td>
          <td>Origin</td>
          <td>Available</td>         
          <td>Freight</td>
          <td>Type</td>
          <td>For sale</td>
          <td>Pre-sold</td>
          <td>Return to</td>
          <td>Display</td>
          <td>F&amp;B</td>
          <td>Invoice</td>
          <td>Pallet</td>
          <td>Loading</td>
          <td>Actioned?</td>
          <td>&nbsp;</td>
        </tr>
        <%= strDisplayList %>
      </table></td>
  </tr>
</table>
</body>
</html>