<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<% strSection = "transfer" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Transfer Requests</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/javascript.js"></script>
<script language="JavaScript" type="text/javascript">
function searchTransfer(){
    var strSearch = document.forms[0].txtSearch.value;
	var strStatus = document.forms[0].cboStatus.value;

    document.location.href = 'list_transfer.asp?type=search&txtSearch=' + strSearch + '&cboStatus=' + strStatus;
}

function resetSearch(){
	document.location.href = 'list_transfer.asp?type=reset';
}
</script>
</head>
<body>
<%
sub setSearch
	select case Trim(Request("type"))
		case "reset"
			session("transfer_search") 			= ""
			session("strSort") 					= ""
			session("transfer_status") 			= 1
			session("transfer_initial_page") 	= 1
		case "search"
			session("transfer_search") 			= Trim(Request("txtSearch"))
			session("strSort") 					= Trim(Request("cboSort"))
			session("transfer_status") 			= Trim(Request("cboStatus"))
			session("transfer_initial_page") 	= 1
	end select
end sub

sub displayTransfer
	dim iRecordCount
	iRecordCount = 0
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

	dim intStatus

	if session("transfer_status") = "" then
		session("transfer_status") = "1"
	end if

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 100

	strSQL = "SELECT * FROM yma_transfer "
	strSQL = strSQL & "	WHERE (	product_1 LIKE '%" & session("transfer_search") & "%' "
	strSQL = strSQL & "			OR product_2 LIKE '%" & session("transfer_search") & "%' "
	strSQL = strSQL & "			OR product_3 LIKE '%" & session("transfer_search") & "%' "
	strSQL = strSQL & "			OR product_4 LIKE '%" & session("transfer_search") & "%' "
	strSQL = strSQL & "			OR product_5 LIKE '%" & session("transfer_search") & "%' "
	strSQL = strSQL & "			OR product_6 LIKE '%" & session("transfer_search") & "%' "
	strSQL = strSQL & "			OR product_7 LIKE '%" & session("transfer_search") & "%' "
	strSQL = strSQL & "			OR product_8 LIKE '%" & session("transfer_search") & "%' "
	strSQL = strSQL & "			OR product_9 LIKE '%" & session("transfer_search") & "%' "
	strSQL = strSQL & "			OR product_10 LIKE '%" & session("transfer_search") & "%' "
	strSQL = strSQL & "			OR product_11 LIKE '%" & session("transfer_search") & "%' "
	strSQL = strSQL & "			OR product_12 LIKE '%" & session("transfer_search") & "%' "
	strSQL = strSQL & "			OR product_13 LIKE '%" & session("transfer_search") & "%' "
	strSQL = strSQL & "			OR product_14 LIKE '%" & session("transfer_search") & "%' "
	strSQL = strSQL & "			OR product_15 LIKE '%" & session("transfer_search") & "%' "
	strSQL = strSQL & "			OR product_16 LIKE '%" & session("transfer_search") & "%' "
	strSQL = strSQL & "			OR product_17 LIKE '%" & session("transfer_search") & "%' "
	strSQL = strSQL & "			OR product_18 LIKE '%" & session("transfer_search") & "%' "
	strSQL = strSQL & "			OR product_19 LIKE '%" & session("transfer_search") & "%' "
	strSQL = strSQL & "			OR product_20 LIKE '%" & session("transfer_search") & "%') "
	'if session("UsrUserName") = "tonyk" then
	'	strSQL = strSQL & "		AND warehouse LIKE '%excel%' "
	'end if
	strSQL = strSQL & "		AND status LIKE '%" & session("transfer_status") & "%' "
	strSQL = strSQL & "	ORDER BY date_created DESC"

	rs.Open strSQL, conn

	'Response.Write strSQL

	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount

	Select Case Request("Action")
	    case "<<"
		    intpage = 1
			session("transfer_initial_page") = intpage
	    case "<"
		    intpage = Request("intpage") - 1
			session("transfer_initial_page") = intpage

			if session("transfer_initial_page") < 1 then session("transfer_initial_page") = 1
	    case ">"
		    intpage = Request("intpage") + 1
			session("transfer_initial_page") = intpage

			if session("transfer_initial_page") > intPageCount then session("transfer_initial_page") = IntPageCount
	    Case ">>"
		    intpage = intPageCount
			session("transfer_initial_page") = intpage
    end select

    strDisplayList = ""

	if not DB_RecSetIsEmpty(rs) Then

	    rs.AbsolutePage = session("transfer_initial_page")

		For intRecord = 1 To rs.PageSize
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

			strDisplayList = strDisplayList & "<td><a href=""update_transfer.asp?id=" & rs("id") & """>" & rs("id") & "</a></td>"
			strDisplayList = strDisplayList & "<td>" & rs("created_by") & "</td>"
			
			
			strDisplayList = strDisplayList & "<td nowrap>" & rs("warehouse")
			if DateDiff("d",rs("date_created"), strTodayDate) = 0 then
				strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=""0"">"
			end if

			if rs("priority") = 1 then
				strDisplayList = strDisplayList & " <img src=""images/icon_priority.gif"" border=""0"">"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("product_1")
			if rs("product_2") <> "" then
				strDisplayList = strDisplayList & ", " & rs("product_2")
			end if
			if rs("product_3") <> "" then
				strDisplayList = strDisplayList & ", " & rs("product_3")
			end if
			if rs("product_4") <> "" then
				strDisplayList = strDisplayList & ", " & rs("product_4")
			end if
			if rs("product_5") <> "" then
				strDisplayList = strDisplayList & ", " & rs("product_5")
			end if
			if rs("product_6") <> "" then
				strDisplayList = strDisplayList & ", " & rs("product_6")
			end if
			if rs("product_7") <> "" then
				strDisplayList = strDisplayList & ", " & rs("product_7")
			end if
			if rs("product_8") <> "" then
				strDisplayList = strDisplayList & ", " & rs("product_8")
			end if
			if rs("product_9") <> "" then
				strDisplayList = strDisplayList & ", " & rs("product_9")
			end if
			if rs("product_10") <> "" then
				strDisplayList = strDisplayList & ", " & rs("product_10")
			end if
			if rs("product_11") <> "" then
				strDisplayList = strDisplayList & ", " & rs("product_11")
			end if
			if rs("product_12") <> "" then
				strDisplayList = strDisplayList & ", " & rs("product_12")
			end if
			if rs("product_13") <> "" then
				strDisplayList = strDisplayList & ", " & rs("product_13")
			end if
			if rs("product_14") <> "" then
				strDisplayList = strDisplayList & ", " & rs("product_14")
			end if
			if rs("product_15") <> "" then
				strDisplayList = strDisplayList & ", " & rs("product_15")
			end if
			if rs("product_16") <> "" then
				strDisplayList = strDisplayList & ", " & rs("product_16")
			end if
			if rs("product_17") <> "" then
				strDisplayList = strDisplayList & ", " & rs("product_17")
			end if
			if rs("product_18") <> "" then
				strDisplayList = strDisplayList & ", " & rs("product_18")
			end if
			if rs("product_19") <> "" then
				strDisplayList = strDisplayList & ", " & rs("product_19")
			end if
			if rs("product_20") <> "" then
				strDisplayList = strDisplayList & ", " & rs("product_20")
			end if
			strDisplayList = strDisplayList & "</td>"

			if IsNull(rs("date_received")) or rs("date_received") = "01/01/1900" or rs("date_received") = "1/1/1900"  then
				strDisplayList = strDisplayList & "<td class=""orange_text"">NA</td>"
			else
				strDisplayList = strDisplayList & "<td nowrap>" & FormatDateTime(rs("date_received"),1) & "</td>"
			end if

			strDisplayList = strDisplayList & "<td>" & rs("transfer_comments") & "</td>"

			strDisplayList = strDisplayList & "<td>"
			if rs("pickup") = 1 then
				strDisplayList = strDisplayList & "<img src=images/tick.gif>"
			else
				strDisplayList = strDisplayList & "<img src=images/cross.gif>"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>"
			if rs("received") = 1 then
				strDisplayList = strDisplayList & "<img src=images/tick.gif>"
			else
				strDisplayList = strDisplayList & "<img src=images/cross.gif>"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("invoice_no") & "</td>"
			strDisplayList = strDisplayList & "<td>"
			if rs("base") = 1 then
				strDisplayList = strDisplayList & "<img src=images/tick.gif>"
			else
				strDisplayList = strDisplayList & "<img src=images/cross.gif>"
			end if
			strDisplayList = strDisplayList & "</td>"
			if rs("status") = 1 then
				strDisplayList = strDisplayList & "<td>Open</td>"
			else
				strDisplayList = strDisplayList & "<td class=""green_text"">Completed</td>"
			end if
			strDisplayList = strDisplayList & "<td><a onclick=""return confirm('Are you sure you want to delete " & rs("id") & " ?');"" href='delete_transfer.asp?id=" & rs("id") & "'><img src=""images/btn_delete.gif"" border=""0""></a></td>"
			strDisplayList = strDisplayList & "</tr>"

			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next
	else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""15"" align=""center"">No transfers found.</td></tr>"
	end if

	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""15"" class=""recordspaging"">"
	strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_transfer.asp"" method=""post"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & session("transfer_initial_page") & ">"

	if session("transfer_initial_page") = 1 then
   		strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
	end if
	if session("transfer_initial_page") = intpagecount or intRecordCount = 0 then
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">>"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">>"">"
	end if
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strSearch & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboState"" value=" & strState & ">"
	strDisplayList = strDisplayList & "<br />"
    strDisplayList = strDisplayList & "Page: " & session("transfer_initial_page") & " to " & intpagecount
    strDisplayList = strDisplayList & "<br />"
	strDisplayList = strDisplayList & "<h2>Search results: " & intRecordCount & " tranfers.</h2>"
    strDisplayList = strDisplayList & "</form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseDataBase()
end sub

sub main
	call UTL_validateLogin
	call setSearch
	
    if trim(session("transfer_initial_page"))  = "" then
    	session("transfer_initial_page") = 1
	end if

    call displayTransfer	
end sub

call main

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList
dim strDealerResultList

%>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><img src="images/icon_transfer.jpg" border="0" alt="Transfer Requests" /></td>
          <td valign="top"><div class="alert alert-success"><img src="images/add_icon.png" border="0" align="bottom" /> <a href="add_transfer.asp">Add Transfer</a></div>
            <p><img src="images/icon_excel.jpg" border="0" /> <a href="export_transfer.asp?search=<%= session("transfer_search") %>&status=<%= session("transfer_status") %>">Export</a></p></td>
          <td valign="top"><div class="alert alert-search">
              <form name="frmSearch" id="frmSearch" action="list_transfer.asp?type=search" method="post" onsubmit="searchTransfer()">
                <h3>Search Parameters:</h3>
                Products:
                <input type="text" name="txtSearch" size="25" value="<%= request("txtSearch") %>" maxlength="20" />
                <select name="cboStatus" onchange="searchTransfer()">
                  <option <% if session("transfer_status") = "1" then Response.Write " selected" end if%> value="1">Open</option>
                  <option <% if session("transfer_status") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
                </select>
                <input type="button" name="btnSearch" value="Search" onclick="searchTransfer()" />
                <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
              </form>
            </div></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td><table cellspacing="0" cellpadding="8" class="database_records">
    <thead>
        <tr>
          <td>ID</td>
          <td nowrap="nowrap">Created</td>
          <td nowrap="nowrap">From - To</td>
          <td>Product(s)</td>
          <td>Received</td>
          <td>Comments</td>
          <td>Picked up?</td>
          <td>Received?</td>
          <td>Invoice</td>
          <td>Base?</td>
          <td>Status</td>
          <td>&nbsp;</td>
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