<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<% strSection = "roadshow" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Roadshow 2014</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/javascript.js"></script>
<script language="JavaScript" type="text/javascript">
function searchItem(){
    var strSearch 			= document.forms[0].txtSearch.value;
	var strItemDepartment  	= document.forms[0].cboDepartment.value;
	var strCategory  		= document.forms[0].cboCategory.value;
	var strItemOwner  		= document.forms[0].cboOwner.value;
	var strSkuType  		= document.forms[0].cboSkuType.value;
	var strItemOrigin  		= document.forms[0].cboOrigin.value;
	var strTransit  		= document.forms[0].cboTransit.value;
	var strReturnTo  		= document.forms[0].cboReturnTo.value;
    //if (strSearch != ''){
     document.location.href = 'list_roadshow.asp?type=search&txtSearch=' + strSearch + '&owner=' + strItemOwner + '&return=' + strReturnTo + '&origin=' + strItemOrigin + '&department=' + strItemDepartment + '&sku=' + strSkuType + '&transit=' + strTransit + '&category=' + strCategory;
    //}
}

function resetSearch(){
	document.location.href = 'list_roadshow.asp?type=reset';
}
</script>
</head>
<body>
<%
sub setSearch
	select case Trim(Request("type"))
		case "reset"
			session("roadshow_Search") 			= ""
			session("roadshow_ItemDepartment") 	= ""
			session("roadshow_ItemOwner") 		= ""
			session("roadshow_Category") 		= ""
			session("roadshow_SkuType") 		= ""
			session("roadshow_ItemOrigin") 		= ""
			session("roadshow_Transit") 		= ""
			session("roadshow_ReturnTo") 		= ""
			session("roadshow_initial_page") 	= 1
		case "search"
			session("roadshow_Search") 			= trim(Request("txtSearch"))
			session("roadshow_ItemDepartment") 	= request("department")
			session("roadshow_ItemOwner") 		= request("owner")
			session("roadshow_Category") 		= request("category")
			session("roadshow_SkuType") 		= request("sku")
			session("roadshow_ItemOrigin") 		= request("origin")
			session("roadshow_Transit") 		= request("transit")
			session("roadshow_ReturnTo") 		= request("return")
			session("roadshow_initial_page") 	= 1
	end select
end sub

sub displayRoadshow
	dim iRecordCount
	iRecordCount = 0
	
    dim strSQL

	dim strPageResultNumber
	dim strRecordPerPage
	dim intRecordCount

	dim strTodayDate
	strTodayDate = FormatDateTime(Date())

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 200

	strSQL = "SELECT * FROM yma_roadshow_2014 "
	strSQL = strSQL & "	WHERE category LIKE '%" & session("roadshow_Category") & "%' "
	strSQL = strSQL & "		AND department LIKE '%" & session("roadshow_ItemDepartment") & "%' "
	strSQL = strSQL & "		AND sku_type LIKE '%" & session("roadshow_SkuType") & "%' "
	strSQL = strSQL & "		AND transit LIKE '%" & session("roadshow_Transit") & "%' "
	strSQL = strSQL & "		AND return_to LIKE '%" & session("roadshow_ReturnTo") & "%' "
	strSQL = strSQL & "		AND origin LIKE '%" & session("roadshow_ItemOrigin") & "%' "
	strSQL = strSQL & "		AND created_by LIKE '%" & session("roadshow_ItemOwner") & "%' "
	strSQL = strSQL & "		AND (product_code LIKE '%" & session("roadshow_Search") & "%' "
	strSQL = strSQL & "			OR description LIKE '%" & session("roadshow_Search") & "%' "
	strSQL = strSQL & "			OR pallet_no LIKE '%" & session("roadshow_Search") & "%' "
	strSQL = strSQL & "			OR loading_sequence LIKE '%" & session("roadshow_Search") & "%') "
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
			if Session("UsrLoginRole") = 1 then
				strDisplayList = strDisplayList & "<td nowrap><a href=""update_roadshow.asp?id=" & rs("item_id") & """>" & rs("item_id") & "</a></td>"
			else			
				strDisplayList = strDisplayList & "<td>" & rs("item_id") & "</td>"
			end if
			strDisplayList = strDisplayList & "<td>" & rs("created_by") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("department") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("item_group") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("category") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("product_code") & "</td>"
			'if DateDiff("d",rs("date_created"), strTodayDate) = 0 then
			'	strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=""0"">"
			'end if
			strDisplayList = strDisplayList & "</td>"
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
			strDisplayList = strDisplayList & "<td>" & rs("how_displayed") & "</td>"
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
			'strDisplayList = strDisplayList & "<td></td>"
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
	strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_roadshow.asp"" method=""post"">"
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
    <td class="first_content"><h2>Roadshow 2014</h2>      
      <!--<p>Final Product List Change: <span style="color:#F00"><strong>TBA</strong></span> - <em>Please note that after this date you will not be able to update the Product List.</em></p>--> 
      <!--<img src="images/forward_arrow.gif" border="0" /> <a href="add_roadshow.asp">Add New Roadshow Item</a>-->
      <p align="right"><img src="images/icon_excel.jpg" border="0" /> <a href="export_roadshow.asp">Export</a></p>      
      <!--<p>All | <a href="list_item-today.asp">Updated Today</a></p>-->      
      <div class="alert alert-search">
        <form name="frmSearch" id="frmSearch" action="list_roadshow.asp?type=search" method="post" onsubmit="searchItem()">
          <h3>Search Parameters:</h3>
          Item / Description / Pallet no / Loading sequence:
          <input type="text" name="txtSearch" size="15" value="<%= request("txtSearch") %>" />
          
          <select name="cboOwner" onchange="searchItem()">
            <option value="">All Owners</option>
            <option <% if session("roadshow_ItemOwner") = "euan" then Response.Write " selected" end if%> value="euan">Euan McInnes</option>
            <option <% if session("roadshow_ItemOwner") = "eric" then Response.Write " selected" end if%> value="eric">Eric Ong</option>
            <option <% if session("roadshow_ItemOwner") = "jamie" then Response.Write " selected" end if%> value="jamie">Jamie Goff</option>
            <option <% if session("roadshow_ItemOwner") = "matt" then Response.Write " selected" end if%> value="matt">Mat Taylor</option>
            <option <% if session("roadshow_ItemOwner") = "mick" then Response.Write " selected" end if%> value="mick">Mick Hughes</option>
            <option <% if session("roadshow_ItemOwner") = "nathan" then Response.Write " selected" end if%> value="nathan">Nathan Biggin</option>
          </select>
          <select name="cboDepartment" onchange="searchItem()">
            <option value="">All Depts</option>
            <option <% if session("roadshow_ItemDepartment") = "ca" then Response.Write " selected" end if%> value="ca">CA</option>
            <option <% if session("roadshow_ItemDepartment") = "pro" then Response.Write " selected" end if%> value="pro">Pro</option>
            <option <% if session("roadshow_ItemDepartment") = "trad" then Response.Write " selected" end if%> value="trad">Trad</option>
          </select>
          <select name="cboCategory" onchange="searchItem()">
            <option value="">All Categories</option>
            <option <% if session("roadshow_Category") = "acoustic" then Response.Write " selected" end if%> value="acoustic">Acoustic Drums</option>
            <option <% if session("roadshow_Category") = "brass" then Response.Write " selected" end if%> value="brass">Brass &amp; Woodwind</option>
            <option <% if session("roadshow_Category") = "ca" then Response.Write " selected" end if%> value="ca">CA</option>
            <option <% if session("roadshow_Category") = "digital" then Response.Write " selected" end if%> value="digital">Digital Pianos</option>
            <option <% if session("roadshow_Category") = "electronic" then Response.Write " selected" end if%> value="electronic">Electronic Drums</option>
            <option <% if session("roadshow_Category") = "guitar" then Response.Write " selected" end if%> value="guitar">Guitars</option>
            <option <% if session("roadshow_Category") = "paiste" then Response.Write " selected" end if%> value="paiste">Paiste</option>
            <option <% if session("roadshow_Category") = "pianos" then Response.Write " selected" end if%> value="pianos">Pianos</option>
            <option <% if session("roadshow_Category") = "portable" then Response.Write " selected" end if%> value="portable">Portable Keyboards</option>
            <option <% if session("roadshow_Category") = "Pro Audio" then Response.Write " selected" end if%> value="Pro Audio">Pro Audio</option>
            <option <% if session("roadshow_Category") = "syde" then Response.Write " selected" end if%> value="syde">SYDE</option>
            <option <% if session("roadshow_Category") = "vox" then Response.Write " selected" end if%> value="vox">VOX</option>
          </select>
          <select name="cboSkuType" onchange="searchItem()">
            <option value="">All SKU</option>
            <option <% if session("roadshow_SkuType") = "A Sku" then Response.Write " selected" end if%> value="A Sku">A Sku</option>
            <option <% if session("roadshow_SkuType") = "B Sku" then Response.Write " selected" end if%> value="B Sku">B Sku</option>
            <option <% if session("roadshow_SkuType") = "New" then Response.Write " selected" end if%> value="New">New</option>
          </select>
          <select name="cboOrigin" onchange="searchItem()">
            <option value="">All Origins</option>
            <option <% if session("roadshow_ItemOrigin") = "3T" then Response.Write " selected" end if%> value="3T">3T</option>
            <option <% if session("roadshow_ItemOrigin") = "Head Office" then Response.Write " selected" end if%> value="Head Office">Head Office</option>
            <option <% if session("roadshow_ItemOrigin") = "YCJ" then Response.Write " selected" end if%> value="YCJ">Direct from YCJ</option>
          </select>
          <select name="cboTransit" onchange="searchItem()">
            <option value="">All Freights</option>
            <option <% if session("roadshow_Transit") = "Sea Freight" then Response.Write " selected" end if%> value="Sea Freight">Sea Freight</option>
            <option <% if session("roadshow_Transit") = "Air Freight" then Response.Write " selected" end if%> value="Air Freight">Air Freight</option>
            <option <% if session("roadshow_Transit") = "NA" then Response.Write " selected" end if%> value="NA">NA</option>
          </select>
          <select name="cboReturnTo" onchange="searchItem()">
            <option value="">Aftershow Return to</option>
            <option <% if session("roadshow_ReturnTo") = "3S" then Response.Write " selected" end if%> value="3S">3S</option>           
            <option <% if session("roadshow_ReturnTo") = "Head Office" then Response.Write " selected" end if%> value="Head Office">Head Office</option>
            <option <% if session("roadshow_ReturnTo") = "Dealer" then Response.Write " selected" end if%> value="Dealer">Dealer</option>
          </select>
          <input type="button" name="btnSearch" value="Search" onclick="searchItem()" />
          <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
        </form>
      </div></td>
  </tr>
  <tr>
    <td><table cellspacing="0" cellpadding="5" class="database_records">
        <tr class="innerdoctitle">          
          <td>ID</td>
          <td>Owner</td>
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
          <td>How displayed</td>
          <td>F&amp;B</td>
          <td>Invoice</td>
          <td>Pallet</td>
          <td>Loading</td>
          <td>Actioned?</td>
        </tr>
        <%= strDisplayList %>
      </table></td>
  </tr>
</table>
</body>
</html>