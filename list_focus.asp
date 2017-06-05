<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<% strSection = "roadshow" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>FOCUS 2015</title>
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
     document.location.href = 'list_focus.asp?type=search&txtSearch=' + strSearch + '&cboReturnTo=' + strReturnTo + '&cboOrigin=' + strItemOrigin + '&cboDepartment=' + strItemDepartment + '&cboCategory=' + strCategory;
    //}
}

function resetSearch(){
	document.location.href = 'list_focus.asp?type=reset';
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
	rs.PageSize = 500

	strSQL = "SELECT * FROM tbl_focus "
	strSQL = strSQL & "	WHERE product_code LIKE '%" & session("strSearch") & "%' "
	strSQL = strSQL & "			OR stock_situation LIKE '%" & session("strSearch") & "%' "
	strSQL = strSQL & "			OR display LIKE '%" & session("strSearch") & "%' "
	strSQL = strSQL & "			OR instruction LIKE '%" & session("strSearch") & "%' "
	strSQL = strSQL & "	ORDER BY id"

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
			
			'if session("UsrUserName") = "johannas" then
				strDisplayList = strDisplayList & "<td nowrap><a href=""update_focus.asp?id=" & rs("id") & """>" & rs("id") & "</a></td>"
			'else
			'	strDisplayList = strDisplayList & "<td></td>"
			'end if				
			strDisplayList = strDisplayList & "<td>" & rs("owner") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("type") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("quantity") & "</td>"			
			strDisplayList = strDisplayList & "<td>" & rs("product_code") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("location") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("stock_situation") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("loan_account") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("display") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("available_for_sale") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("instruction") & "</td>"	
			strDisplayList = strDisplayList & "<td>" & rs("notes") & "</td>"			
			strDisplayList = strDisplayList & "<td>"
			if rs("logistics_action") = 1 then
				strDisplayList = strDisplayList & "Yes"
			else
				strDisplayList = strDisplayList & "No"
			end if
			strDisplayList = strDisplayList & "</td>"	
			strDisplayList = strDisplayList & "<td>" & rs("invoice_no") & "</td>"		
			strDisplayList = strDisplayList & "<td>" & rs("pallet_no") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("loading_sequence") & "</td>"
			strDisplayList = strDisplayList & "<td><a onclick=""return confirm('Are you sure you want to delete " & rs("id") & " ?');"" href='delete_focus.asp?id=" & rs("id") & "'><img src=""images/btn_delete.gif"" border=""0""></a></td>"
			strDisplayList = strDisplayList & "</tr>"

			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next

	else
        strDisplayList = "<tr class=innerdoc><td colspan=""16"" align=""center"">No items found.</td></tr>"
	end if

	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""16"" class=""recordspaging"">"
	strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_focus.asp"" method=""post"">"
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

	'response.Write session("roadshow_initial_page")
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
    <td class="first_content"><h2>FOCUS 2015</h2>
      <p><a href="export_focus.asp">Export this</a></p>
      <!--<p align="right"><img src="images/icon_excel.jpg" border="0" /> <a href="export_focus.asp">Export all</a></p>-->
      <div class="alert alert-search">
        <form name="frmSearch" id="frmSearch" action="list_focus.asp?type=search" method="post" onsubmit="searchItem()">
          <h3>Search:</h3>         
          <input type="text" name="txtSearch" size="30" maxlength="30" value="<%= request("txtSearch") %>" />                   
          <input type="button" name="btnSearch" value="Search" onclick="searchItem()" />
          <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
        </form>
      </div></td>
  </tr>
  <tr>
    <td><table cellspacing="0" cellpadding="4" class="database_records">
        <tr class="innerdoctitle">          
          <td>ID</td>
          <td>Owner</td>
          <td>Type</td>
          <td>Qty</td>
          <td>Product Code</td>
          <td>Location</td>
          <td>Stock Situation</td>
          <td>Loan Account</td>
          <td>Display</td>
          <td>For sale</td>          
          <td>Instruction</td>
          <td>Notes</td>          
          <td>Actioned?</td>  
          <td>Invoice</td>  
          <td>Pallet</td>  
          <td>Loading sequence</td>   
          <td></td>       
        </tr>
        <%= strDisplayList %>
      </table></td>
  </tr>
</table>
</body>
</html>