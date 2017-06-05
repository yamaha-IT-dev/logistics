<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<% strSection = "cancelled" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Cancelled Orders</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/javascript.js"></script>
<script language="JavaScript" type="text/javascript">
function searchCancelled(){
    var strSearch = document.forms[0].txtSearch.value;
	var strStatus  = document.forms[0].cboStatus.value;
    document.location.href = 'list_cancelled.asp?type=search&txtSearch=' + strSearch + '&cboStatus=' + strStatus;
}

function resetSearch(){
	document.location.href = 'list_cancelled.asp?type=reset';
}
</script>
</head>
<body>
<%
sub setSearch
	select case Trim(Request("type"))
		case "reset"
			session("cancelled_search") 		= ""
			session("cancelled_status") 		= ""
			session("cancelled_initial_page") 	= 1
		case "search"
			session("cancelled_search") 		= Trim(Request("txtSearch"))
			session("cancelled_status") 		= Trim(Request("cboStatus"))
			session("cancelled_initial_page") 	= 1
	end select
end sub

sub displayCancelled
	dim iRecordCount
	iRecordCount = 0
    Dim strSortBy
	dim strSortItem
    Dim strSearch
    dim strSQL
	dim strType
	dim strSort
	dim strStatus
	
	dim strRecordPerPage
	dim intRecordCount
	dim strModifiedDate

	dim strTodayDate
	strTodayDate = FormatDateTime(Date())

	if session("cancelled_status") = "" then
		session("cancelled_status") = "1"
	end if

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 100

	strSQL = "SELECT * FROM yma_cancelled_order "
	strSQL = strSQL & "	WHERE  (cancel_shipment_no LIKE '%" & session("cancelled_search") & "%' "
	strSQL = strSQL & "			OR cancel_info LIKE '%" & session("cancelled_search") & "%' "
	strSQL = strSQL & "			OR cancel_created_by LIKE '%" & session("cancelled_search") & "%') "
	strSQL = strSQL & "		AND cancel_status LIKE '%" & session("cancelled_status") & "%' "
	strSQL = strSQL & "	ORDER BY cancel_date_created DESC"

	rs.Open strSQL, conn

	'Response.Write strSQL

	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount

	Select Case Request("Action")
	    case "<<"
		    intpage = 1
			session("cancelled_initial_page") = intpage
	    case "<"
		    intpage = Request("intpage") - 1
			session("cancelled_initial_page") = intpage

			if session("cancelled_initial_page") < 1 then session("cancelled_initial_page") = 1
	    case ">"
		    intpage = Request("intpage") + 1
			session("cancelled_initial_page") = intpage

			if session("cancelled_initial_page") > intPageCount then session("cancelled_initial_page") = IntPageCount
	    Case ">>"
		    intpage = intPageCount
			session("cancelled_initial_page") = intpage
    end select

    strDisplayList = ""

	if not DB_RecSetIsEmpty(rs) Then

	    rs.AbsolutePage = session("cancelled_initial_page")

		For intRecord = 1 To rs.PageSize
			if (DateDiff("d",rs("cancel_date_modified"), strTodayDate) = 0) OR (DateDiff("d",rs("cancel_date_created"), strTodayDate) = 0) then
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

			strDisplayList = strDisplayList & "<td><a href=""update_cancelled_order.asp?id=" & rs("cancel_id") & """><img src=""images/icon_view.png"" border=""0""></a></td>"
			strDisplayList = strDisplayList & "<td>" & Lcase(rs("cancel_created_by")) & "</td>"
			strDisplayList = strDisplayList & "<td>" & WeekDayName(WeekDay(rs("cancel_date_created"))) & ", " & FormatDateTime(rs("cancel_date_created"),1) & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("cancel_shipment_no")
			if DateDiff("d",rs("cancel_date_created"), strTodayDate) = 0 then
				strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=""0"">"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("cancel_info") & "</td>"	
			strDisplayList = strDisplayList & "<td>"				
			Select Case	rs("cancel_warehouse_confirm")
				case 1
					strDisplayList = strDisplayList & "<img src=images/tick.gif>"
				case 0
					strDisplayList = strDisplayList & "<img src=images/cross.gif>"
				case else
					strDisplayList = strDisplayList & "Not yet"
			end select				
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>"				
			Select Case	rs("cancel_logistics_confirm")
				case 1
					strDisplayList = strDisplayList & "<img src=images/tick.gif>"
				case 0
					strDisplayList = strDisplayList & "<img src=images/cross.gif>"
				case else
					strDisplayList = strDisplayList & "Not yet"
			end select				
			strDisplayList = strDisplayList & "</td>"
			
			select case rs("cancel_status") 
				case 1
					strDisplayList = strDisplayList & "<td>Open"
				case 2
					strDisplayList = strDisplayList & "<td align=""orange_text"">Cancelled"
				case else
					strDisplayList = strDisplayList & "<td class=""green_text"">Completed"
			end select
			strDisplayList = strDisplayList & "</td>"
			
			strDisplayList = strDisplayList & "<td>" & rs("cancel_modified_by") & ""
			
			if IsNull(rs("cancel_modified_by")) then
				strDisplayList = strDisplayList & "NA</td>"
			else
				strDisplayList = strDisplayList & " - " & WeekDayName(WeekDay(rs("cancel_date_modified"))) & ", " & FormatDateTime(rs("cancel_date_modified"),1) & "</td>"
			end if

			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td><a onclick=""return confirm('Are you sure you want to delete " & rs("cancel_shipment_no") & " ?');"" href='delete_cancelled_order.asp?id=" & rs("cancel_id") & "'><img src=""images/btn_delete.gif"" border=""0""></a></td>"
			strDisplayList = strDisplayList & "</tr>"

			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next

	else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""22"" align=""center"">No cancelled orders found.</td></tr>"
	end if

	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""22"" class=""recordspaging"">"
	strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_cancelled.asp"" method=""post"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & session("cancelled_initial_page") & ">"

	if session("cancelled_initial_page") = 1 then
   		strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
	end if
	if session("cancelled_initial_page") = intpagecount or intRecordCount = 0 then
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">>"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">>"">"
	end if
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strSearch & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboState"" value=" & strState & ">"
	strDisplayList = strDisplayList & "<br />"
    strDisplayList = strDisplayList & "Page: " & session("cancelled_initial_page") & " to " & intpagecount
    strDisplayList = strDisplayList & "<br />"
	strDisplayList = strDisplayList & "<h2>Search results: " & intRecordCount & " orders.</h2>"
    strDisplayList = strDisplayList & "</form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseDataBase()
end sub

sub main
	call UTL_validateLogin
	call setSearch

    if trim(session("cancelled_initial_page"))  = "" then
    	session("cancelled_initial_page") = 1
	end if

    call displayCancelled
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
          <td valign="top"><img src="images/icon_cancelled.jpg" border="0" alt="Cancelled Order" /></td>
          <td valign="top"><div class="alert alert-success"><img src="images/add_icon.png" border="0" align="bottom" /> <a href="../request/add_cancelled_order.asp">Add Cancelled Order</a></div>
            <p><img src="images/icon_excel.jpg" border="0" /> <a href="export_cancelled.asp?search=<%= session("cancelled_search") %>&status=<%= session("cancelled_status") %>">Export</a></p></td>
          <td valign="top"><div class="alert alert-search">
              <form name="frmSearch" id="frmSearch" action="list_cancelled.asp?type=search" method="post" onsubmit="searchCancelled()">
                <h3>Search Parameters:</h3>
                Shipment no / Reason / Created by:
                <input type="text" name="txtSearch" size="25" value="<%= request("txtSearch") %>" maxlength="20" />
                <select name="cboStatus" onchange="searchCancelled()">
                  <option <% if session("cancelled_status") = "1" then Response.Write " selected" end if%> value="1">Open</option>
                  <option <% if session("cancelled_status") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
                </select>
                <input type="button" name="btnSearch" value="Search" onclick="searchCancelled()" />
                <input type="button" name="btnReset" value="Reset" onclick="resetSearch()" />
              </form>
            </div></td>
        </tr>
      </table>
      <table cellspacing="0" cellpadding="8" class="database_records">
      <thead>
        <tr>
          <td width="5%">&nbsp;</td>
          <td width="10%">Created by</td>
          <td width="15%">Date created</td>
          <td width="10%">Shipment no</td>
          <td width="13%">Reason</td>
          <td width="7%">WH confirm</td>
          <td width="10%">Logistics confirm</td>
          <td width="5%">Status</td>        
          <td width="20%">Last modified</td>
          <td width="5%">&nbsp;</td>
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