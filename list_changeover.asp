<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<% strSection = "changeover" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Changeover Logs</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/javascript.js"></script>
<script language="JavaScript" type="text/javascript">
function searchChangeover(){
    var strSearch = document.forms[0].txtSearch.value;
	var strState  = document.forms[0].cboState.value;
	var strStatus  = document.forms[0].cboStatus.value;
    document.location.href = 'list_changeover.asp?type=search&txtSearch=' + strSearch + '&cboState=' + strState + '&cboStatus=' + strStatus;
}

function resetSearch(){
	document.location.href = 'list_changeover.asp?type=reset';
}
</script>
</head>
<body>
<%
sub setSearch
	select case Trim(Request("type"))
		case "reset"
			session("changeover_search") 		= ""
			session("changeover_state") 		= ""
			session("changeover_status") 		= ""
			session("changeover_initial_page") 	= 1
		case "search"
			session("changeover_search") 		= Trim(Request("txtSearch"))
			session("changeover_state") 		= Trim(Request("cboState"))
			session("changeover_status") 		= Trim(Request("cboStatus"))
			session("changeover_initial_page") 	= 1
	end select
end sub

sub displayChangeover
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

	if session("changeover_status") = "" then
		session("changeover_status") = "1"
	end if

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 100

	strSQL = "SELECT * FROM yma_changeover "
	strSQL = strSQL & "	WHERE state LIKE '%" & session("changeover_state") & "%' "
	strSQL = strSQL & "		AND (customer LIKE '%" & session("changeover_search") & "%' "
	strSQL = strSQL & "			OR contact_person LIKE '%" & session("changeover_search") & "%' "
	strSQL = strSQL & "			OR old_model LIKE '%" & session("changeover_search") & "%' "
	strSQL = strSQL & "			OR old_model_serial LIKE '%" & session("changeover_search") & "%' "
	strSQL = strSQL & "			OR invoice_no LIKE '%" & session("changeover_search") & "%' "
	strSQL = strSQL & "			OR connote LIKE '%" & session("changeover_search") & "%') "
	strSQL = strSQL & "		AND status LIKE '%" & session("changeover_status") & "%' "
	strSQL = strSQL & "	ORDER BY changeover_id DESC"

	rs.Open strSQL, conn

	'Response.Write strSQL

	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount

	Select Case Request("Action")
	    case "<<"
		    intpage = 1
			session("changeover_initial_page") = intpage
	    case "<"
		    intpage = Request("intpage") - 1
			session("changeover_initial_page") = intpage

			if session("changeover_initial_page") < 1 then session("changeover_initial_page") = 1
	    case ">"
		    intpage = Request("intpage") + 1
			session("changeover_initial_page") = intpage

			if session("changeover_initial_page") > intPageCount then session("changeover_initial_page") = IntPageCount
	    Case ">>"
		    intpage = intPageCount
			session("changeover_initial_page") = intpage
    end select

    strDisplayList = ""

	if not DB_RecSetIsEmpty(rs) Then

	    rs.AbsolutePage = session("changeover_initial_page")

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

			strDisplayList = strDisplayList & "<td><a href=""update_changeover.asp?id=" & rs("changeover_id") & """>" & rs("changeover_id") & "</a></td>"
			strDisplayList = strDisplayList & "<td><strong>" & rs("customer") & "</strong>"
			if DateDiff("d",rs("date_created"), strTodayDate) = 0 then
				strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=""0"">"
			end if
			strDisplayList = strDisplayList & " (" & rs("contact_person") & ")</td>"
			strDisplayList = strDisplayList & "<td nowrap>" & rs("phone") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("city") & ", " & rs("state") & "</td>"
			strDisplayList = strDisplayList & "<td><strong>" & rs("old_model") & "</strong></td>"
			strDisplayList = strDisplayList & "<td>" & rs("old_model_serial") & "</td>"
			strDisplayList = strDisplayList & "<td>"
			if rs("proof") = 1 then
				strDisplayList = strDisplayList & "<img src=images/tick.gif>"
			else
				strDisplayList = strDisplayList & "<img src=images/cross.gif>"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>"
			if rs("warranty") = 1 then
				strDisplayList = strDisplayList & "<img src=images/tick.gif>"
			else
				strDisplayList = strDisplayList & "<img src=images/cross.gif>"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>"
			if rs("destroy") = 1 then
				strDisplayList = strDisplayList & "<img src=images/tick.gif>"
			else
				strDisplayList = strDisplayList & "<img src=images/cross.gif>"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("replacement_model") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("make_up_cost") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("replacement_destination") & "</td>"
			if IsNull(rs("date_received")) or rs("date_received") = "01/01/1900" or rs("date_received") = "1/1/1900" or rs("date_received") = "01/01/2000" or rs("date_received") = "1/1/2000" then
				strDisplayList = strDisplayList & "<td class=""orange_text"">NA</td>"
			else
				strDisplayList = strDisplayList & "<td>" & rs("date_received") & "</td>"
			end if
			
			if rs("date_payment") = "01/01/1900" or rs("date_payment") = "1/1/1900" or rs("date_payment") = "01/01/2000" or rs("date_payment") = "1/1/2000" then
				strDisplayList = strDisplayList & "<td class=""orange_text"">NA</td>"
			else
				strDisplayList = strDisplayList & "<td>" & rs("date_payment") & "</td>"
			end if

			strDisplayList = strDisplayList & "<td>" & rs("invoice_no") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("connote") & "</td>"

			if rs("status") = 1 then
				strDisplayList = strDisplayList & "<td class=""blue_text"">Open</td>"
			else
				strDisplayList = strDisplayList & "<td class=""green_text"">Completed</td>"
			end if
			strDisplayList = strDisplayList & "<td><a onclick=""return confirm('Are you sure you want to delete " & rs("customer") & " ?');"" href='delete_changeover.asp?id=" & rs("changeover_id") & "'><img src=""images/btn_delete.gif"" border=""0""></a></td>"
			strDisplayList = strDisplayList & "</tr>"

			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next

	else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""18"" align=""center"">No changeovers found.</td></tr>"
	end if

	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""18"" class=""recordspaging"">"
	strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_changeover.asp"" method=""post"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & session("changeover_initial_page") & ">"

	if session("changeover_initial_page") = 1 then
   		strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
	end if
	if session("changeover_initial_page") = intpagecount or intRecordCount = 0 then
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">>"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">>"">"
	end if
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strSearch & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboState"" value=" & strState & ">"
	strDisplayList = strDisplayList & "<br />"
    strDisplayList = strDisplayList & "Page: " & session("changeover_initial_page") & " to " & intpagecount
    strDisplayList = strDisplayList & "<br />"
	strDisplayList = strDisplayList & "<h2>Search results: " & intRecordCount & " changeovers.</h2>"
    strDisplayList = strDisplayList & "</form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseDataBase()
end sub

sub main
	call UTL_validateLogin
	call setSearch

    if trim(session("changeover_initial_page"))  = "" then
    	session("changeover_initial_page") = 1
	end if

    call displayChangeover

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

    strStateList = strStateList & "<option value=''>All states</option>"

    ' We check if there is anything
    if isarray(arrStateFillID) then
        if ubound(arrStateFillID) > 0 then

            for intCounter = 0 to ubound(arrStateFillID)

                if trim(session("changeover_state")) = trim(arrStateFillID(intCounter)) then
                    strStateList = strStateList & "<option selected value=" & arrStateFillID(intCounter) & ">" & arrStateFillText(intCounter) & "</option>"
                else
                   	strStateList = strStateList & "<option value=" & arrStateFillID(intCounter) & ">" & arrStateFillText(intCounter) & "</option>"
                end if

            next
        end if

    end if

end function
%>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><img src="images/icon_changeover.jpg" border="0" alt="Changeover Logs" /></td>
          <td valign="top"><div class="alert alert-success"><img src="images/add_icon.png" border="0" align="bottom" /> <a href="add_changeover.asp">Add Changeover</a></div>
            <p><img src="images/icon_excel.jpg" border="0" /> <a href="export_changeoverlogs.asp?search=<%= session("changeover_search") %>&state=<%= session("changeover_state") %>&status=<%= session("changeover_status") %>">Export</a></p></td>
          <td valign="top"><div class="alert alert-search">
              <form name="frmSearch" id="frmSearch" action="list_changeover.asp?type=search" method="post" onsubmit="searchChangeover()">
                <h3>Search Parameters:</h3>
                Customer / Contact person / Old model / Serial no / Invoice no / Con-note:
                <input type="text" name="txtSearch" size="25" value="<%= request("txtSearch") %>" maxlength="20" />
                <select name="cboState" onchange="searchChangeover()">
                  <%= strStateList %>
                </select>
                <select name="cboStatus" onchange="searchChangeover()">
                  <option <% if session("changeover_status") = "1" then Response.Write " selected" end if%> value="1">Open</option>
                  <option <% if session("changeover_status") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
                </select>
                <input type="button" name="btnSearch" value="Search" onclick="searchChangeover()" />
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
          <td>Customer (Contact person)</td>
          <td>Phone</td>
          <td>Suburb</td>          
          <td>Model No</td>
          <td>Serial No</td>
          <td>Proof</td>
          <td>Warranty</td>
          <td>Destroy</td>
          <td>Replacement</td>
          <td>Cost</td>
          <td>Replacement going to</td>
          <td>Received</td>
          <td>Paid</td>
          <td>Invoice</td>
          <td>Connote</td>
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