<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<% strSection = "freight" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Freight</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/javascript.js"></script>
<script language="JavaScript" type="text/javascript">
function searchFreight(){
    var strSearch = document.forms[0].txtSearch.value;
	var strState  = document.forms[0].cboState.value;
	var strStatus = document.forms[0].cboStatus.value;
	//var strSort = document.forms[0].cboSort.value;
    document.location.href = 'list_freight.asp?type=search&txtSearch=' + strSearch + '&cboState=' + strState + '&cboStatus=' + strStatus;
}

function resetSearch(){
	document.location.href = 'list_freight.asp?type=reset';
}
</script>
</head>
<body>
<%
sub setSearch
	select case Trim(Request("type"))
		case "reset"
			session("freight_search") 	= ""
			session("freight_state") 	= ""
			session("strSort") 			= ""
			session("freight_status") 	= ""
			'session("freight_supplier") = ""
			'session("initialPage") = 1
		case "search"
			session("freight_search") 	= Trim(Request("txtSearch"))
			session("freight_state") 	= Trim(Request("cboState"))
			session("strSort") 			= Trim(Request("cboSort"))
			session("freight_status") 	= Trim(Request("cboStatus"))
			'session("freight_supplier") = Trim(Request("cboSupplier"))
			'session("initialPage") = 1
	end select
end sub

sub displayFreight
	dim strRequester
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

	if session("freight_status") = "" then
		session("freight_status") = "1"
	end if

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic

	strPageResultNumber = trim(Request("cboDealerResultSize"))
	strRecordPerPage = 900

	rs.PageSize = 900

	if Lcase(session("UsrUserName")) = "katet" or Lcase(session("UsrUserName")) = "naomih" then
		strSQL = "SELECT * FROM yma_freight WHERE supplier = 'cope' AND receiver_state LIKE '%" & session("freight_state") & "%' AND (email LIKE '%" & session("freight_search") & "%' OR receiver_name LIKE '%" & session("freight_search") & "%' OR receiver_address LIKE '%" & session("freight_search") & "%' OR items LIKE '%" & session("freight_search") & "%') AND status LIKE '%" & session("freight_status") & "%' ORDER BY date_created DESC"
	else
		strSQL = "SELECT * FROM yma_freight WHERE receiver_state LIKE '%" & session("freight_state") & "%' AND (email LIKE '%" & session("freight_search") & "%' OR receiver_name LIKE '%" & session("freight_search") & "%' OR receiver_address LIKE '%" & session("freight_search") & "%' OR items LIKE '%" & session("freight_search") & "%') AND status LIKE '%" & session("freight_status") & "%' ORDER BY date_created DESC"
	end if

	rs.Open strSQL, conn

	'Response.Write strSQL

	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount

    strDisplayList = ""

	if not DB_RecSetIsEmpty(rs) Then

	    'rs.AbsolutePage = session("cinitialPage")

		For intRecord = 1 To rs.PageSize
			'strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
			strRequester = LCase(rs("email"))
			'strRequester = Split(rs("email"),"@")

			if (DateDiff("d",rs("date_modified"), strTodayDate) = 0) OR (DateDiff("d",rs("date_created"), strTodayDate) = 0) then
				strDisplayList = strDisplayList & "<tr class=""updated_today"">"
			else
				strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
			end if

			strDisplayList = strDisplayList & "<td align=""center""><a href=""update_freight.asp?id=" & rs("id") & """>Edit</a></td>"
			strDisplayList = strDisplayList & "<td align=""center"">"
			if strRequester = "jaclyn_williams@gmx.yamaha.com" then
				strDisplayList = strDisplayList & " <font color=red>" & strRequester & "</font>"
			else
				strDisplayList = strDisplayList & "" & strRequester & ""
			end if

			if DateDiff("d",rs("date_created"), strTodayDate) = 0 then
				strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=""0"">"
			end if

			if rs("priority") = 1 then
				strDisplayList = strDisplayList & " <img src=""images/icon_priority.gif"" border=""0"">"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("supplier") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("pickup_name") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("pickup_city") & " " & rs("pickup_postcode") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("pickup_date") & " - " & rs("pickup_time") & "</td>"
			if rs("return_pickup") = 1 then
				strDisplayList = strDisplayList & "<td align=""center"" bgcolor=""#FFFF00""><img src=images/tick.gif>"
			else
				strDisplayList = strDisplayList & "<td align=""center""><img src=images/cross.gif>"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center""><a href=""http://59.167.253.200/track.php?consignment=" & trim(rs("return_connote")) & """ target=""_blank"">" & trim(rs("return_connote")) & "</a></td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("receiver_name") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("receiver_city") & " " & rs("receiver_postcode") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">" & rs("delivery_date") & " - " & rs("delivery_time") & "</td>"
			strDisplayList = strDisplayList & "<td align=""center"">"
			if rs("pickup") = 1 then
				strDisplayList = strDisplayList & "<img src=images/tick.gif>"
			else
				strDisplayList = strDisplayList & "<img src=images/cross.gif>"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td align=""center""><a href=""http://59.167.253.200/track.php?consignment=" & trim(rs("consignment_no")) & """ target=""_blank"">" & trim(rs("consignment_no")) & "</a></td>"

			if rs("status") = 1 then
				strDisplayList = strDisplayList & "<td align=""center"">Open</td>"
			else
				strDisplayList = strDisplayList & "<td class=""green_text"">Completed</td>"
			end if
			strDisplayList = strDisplayList & "<td align=""center""><a onclick=""return confirm('Are you sure you want to delete " & rs("id") & " ?');"" href='delete_freight.asp?id=" & rs("id") & "'><img src=""images/btn_delete.gif"" border=""0""></a></td>"
			strDisplayList = strDisplayList & "</tr>"

			rs.movenext

			If rs.EOF Then Exit For
		next
	else
        strDisplayList = "<tr class=innerdoc><td colspan=15 align=center>There is no open freight.</td></tr>"
	end if

	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""15"" class=""recordspaging"">"
	'strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_stockmod.asp"" method=""post"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strSearch & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboState"" value=" & strState & ">"
	'strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboSort"" value=" & strSort & ">"
	'strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSource"" value=" & strSource & ">"
    strDisplayList = strDisplayList & "<br />"
	strDisplayList = strDisplayList & "Search results: " & intRecordCount & " records."
    'strDisplayList = strDisplayList & "</form>"
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

    call displayFreight

	call getStateList
end sub

call main

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList
dim strDealerResultList
dim strStateList

'-----------------------------------------------
'				GET STATES
'-----------------------------------------------
function getStateList
    dim arrStateFillText
    dim arrStateFillID
    dim intCounter

    arrStateFillText        = split(arrStateText, ",")
    arrStateFillID 		    = split(arrStateID, ",")

    strStateList = strStateList & "<option value=''>All States</option>"

    ' We check if there is anything
    if isarray(arrStateFillID) then
        if ubound(arrStateFillID) > 0 then

            for intCounter = 0 to ubound(arrStateFillID)

                if trim(session("freight_state")) = trim(arrStateFillID(intCounter)) then
                    strStateList = strStateList & "<option selected value=" & arrStateFillID(intCounter) & ">" & arrStateFillText(intCounter) & "</option>"
                else
                   	strStateList = strStateList & "<option value=" & arrStateFillID(intCounter) & ">" & arrStateFillText(intCounter) & "</option>"
                end if

            next
        end if

    end if

end function
%>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table">
<!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content">
    <table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td><img src="images/icon_freight.jpg" border="0" alt="Freight Requests" /></td>
          <td valign="top">
          <img src="images/icon_excel.jpg" border="0" /> <a href="export_freight.asp?search=<%= session("freight_search") %>&state=<%= session("freight_state") %>&status=<%= session("freight_status") %>">Export</a></td>
        </tr>
      </table>
      <form name="frmSearch" id="frmSearch" action="list_freight.asp?type=search" method="post" onsubmit="searchFreight()">
        <p>Search Requested By / Receiver Name / Address / Items:
          <input type="text" name="txtSearch" size="15" value="<%= request("txtSearch") %>" />
          <select name="cboState">
          	<%= strStateList %>
          </select>
          <select name="cboStatus">
            <option <% if session("freight_status") = "1" then Response.Write " selected" end if%> value="1">Open</option>
            <option <% if session("freight_status") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
          </select>
          <input type="button" name="btnSearch" value="Search" onclick="searchFreight()" />
          <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
      </form></td>
  </tr>
  <tr>
    <td class="database_column"><table cellspacing="0" cellpadding="4" class="database_records">
        <tr class="innerdoctitle" align="center">
          <td>&nbsp;</td>
          <td>Requested by</td>
          <td>Supplier</td>
          <td>Pickup</td>
          <td>Location</td>
          <td>Date / time</td>
          <td bgcolor="#FF0000">Return to pickup?</td>
          <td>Return Connote</td>
          <td bgcolor="#666666">Receiver</td>
          <td bgcolor="#666666">Location</td>
          <td bgcolor="#666666">Date / time</td>
          <td>Picked up?</td>
          <td>Connote</td>
          <td>Status</td>
          <td>&nbsp;</td>
        </tr>
        <%= strDisplayList %>
      </table></td>
  </tr>
</table>
</body>
</html>