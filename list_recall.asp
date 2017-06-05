<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<% strSection = "recall" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Customer Recall</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/javascript.js"></script>
<script language="JavaScript" type="text/javascript">
function searchRecall(){
    var strSearch = document.forms[0].txtSearch.value;
    document.location.href = 'list_recall.asp?type=search&txtSearch=' + strSearch;
}

function resetSearch(){
	document.location.href = 'list_recall.asp?type=reset';
}
</script>
</head>
<body>
<%
sub setSearch
	select case Trim(Request("type"))
		case "reset"
			session("strSearch") 	= ""
			session("strState") 	= ""
			session("strSort") 		= ""
			session("strStatus") 	= ""
		case "search"
			session("strSearch") 	= Trim(Request("txtSearch"))
			session("strState") 	= Trim(Request("cboState"))
			session("strSort") 		= Trim(Request("cboSort"))
			session("strStatus") 	= Trim(Request("cboStatus"))
	end select
end sub

sub displayRecall

    Dim strSortBy
	dim strSortItem
    Dim strSearch
    dim strSQL
	dim strType
	dim strSort
	dim strStatus

	dim strPageResultNumber
	dim strRecordPerPage
	dim intRecordCount
	dim strModifiedDate

	dim strTodayDate
	strTodayDate = FormatDateTime(Date())

	strState = ""
    if len(trim(Session("strState"))) > 0 then
        strType = Session("strState")
    end if

	strSort = ""
    if len(trim(Session("strSort"))) > 0 then
        strSort = Session("strSort")
    end if

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strPageResultNumber = trim(Request("cboDealerResultSize"))
	strRecordPerPage = 900

	rs.PageSize = 900

	strSQL = "SELECT * FROM yma_customer_recall WHERE (customer_name LIKE '%" & Trim(Request("txtSearch")) & "%' OR dealer LIKE '%" & Trim(Request("txtSearch")) & "%' OR product LIKE '%" & Trim(Request("txtSearch")) & "%') ORDER BY recall_id"

	rs.Open strSQL, conn

	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount

    strDisplayList = ""

	if not DB_RecSetIsEmpty(rs) Then

		For intRecord = 1 To rs.PageSize
			if (DateDiff("d",rs("date_modified"), strTodayDate) = 0) OR (DateDiff("d",rs("date_created"), strTodayDate) = 0) then
				strDisplayList = strDisplayList & "<tr class=""updated_today"">"
			else
				strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
			end if

			strDisplayList = strDisplayList & "<td align=""center""><a href=""update_recall.asp?id=" & rs("recall_id") & """>Edit</a></td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("recall_id") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("dealer") & ""
			if DateDiff("d",rs("date_created"), strTodayDate) = 0 then
				strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=""0"">"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("product") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("qty") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("customer_name") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("customer_address") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("customer_city") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("customer_state") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("customer_postcode") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("customer_email") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("customer_phone") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("customer_mobile") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("comments") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center""><a onclick=""return confirm('Are you sure you want to delete " & rs("customer_name") & " ?');"" href='delete_recall.asp?id=" & rs("recall_id") & "'><img src=""images/btn_delete.gif"" border=""0""></a></td>"
			strDisplayList = strDisplayList & "</tr>"

			rs.movenext

			If rs.EOF Then Exit For
		next

	else
        strDisplayList = "<tr class=innerdoc><td colspan=14 align=center>There is no customer recall records.</td></tr>"
	end if

	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""14"" class=""recordspaging"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strSearch & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboState"" value=" & strState & ">"
    strDisplayList = strDisplayList & "<br />"
	strDisplayList = strDisplayList & "Search results: " & intRecordCount & " records."
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

	'Set rs = nothing
    call CloseDataBase()
end sub

sub main
	call UTL_validateLogin
	call setSearch
	'Response.Write "<br>Previous: " & Request.ServerVariables("HTTP_REFERER") & Request.Querystring
    if trim(session("cinitialPage"))  = "" then
    	session("cinitialPage") = 1
	end if

    call displayRecall

	call getStateList
end sub

call main

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList
%>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table">
<!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content">
    <h2>Customer Recall</h2>
    <img src="images/forward_arrow.gif" border="0" /> <a href="add_recall.asp">Add NEW Customer Recall</a>
    <p align="right"><img src="images/icon_excel.jpg" border="0" /> <a href="export_recall.asp">Export ALL</a></p>
      <form name="frmSearch" id="frmSearch">
        <p>Search Dealer / Product / Customer name:
          <input type="text" name="txtSearch" size="15" value="<%= request("txtSearch") %>" />
          <input type="button" name="btnSearch" value="Search" onclick="searchRecall()" />
          <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
      </p>
    </form></td>
  </tr>
  <tr>
    <td class="database_column"><table cellspacing="0" cellpadding="4" class="database_records">
        <tr class="innerdoctitle" align="center">
          <td>&nbsp;</td>
          <td>ID</td>
          <td>Dealer</td>
          <td>Product</td>
          <td>Qty</td>
          <td width="15%">Customer</td>
          <td>Address</td>
          <td>City</td>
          <td>State</td>
          <td>Postcode</td>
          <td>Email</td>
          <td>Phone</td>
          <td>Mobile</td>
          <td width="15%">Comments</td>
          <td>&nbsp;</td>
        </tr>
        <%= strDisplayList %>
      </table></td>
  </tr>
</table>
</body>
</html>