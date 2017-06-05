<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<% strSection = "freight" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Freight Requests</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/javascript.js"></script>
<script type="text/javascript" src="include/jquery.js"></script>
<script type="text/javascript" src="include/main.js"></script>
<script language="JavaScript" type="text/javascript">
function searchFreight(){
    var strSearch = document.forms[0].txtSearch.value;
	var strState  = document.forms[0].cboState.value;
	var strStatus = document.forms[0].cboStatus.value;
	var strEmail = document.forms[0].cboEmail.value;
	
    document.location.href = 'list_freights.asp?type=search&txtSearch=' + strSearch + '&cboState=' + strState + '&cboEmail=' + strEmail  + '&cboStatus=' + strStatus;
}

function resetSearch(){
	document.location.href = 'list_freights.asp?type=reset';
}
</script>
</head>
<body>
<%
sub setSearch
	select case Trim(Request("type"))
		case "reset"
			session("freight_search") 		= ""
			session("freight_state") 		= ""
			session("freight_email") 		= ""
			session("freight_status") 		= ""			
			session("freight_initial_page") = 1
		case "search"
			session("freight_search") 		= Trim(Request("txtSearch"))
			session("freight_state") 		= Trim(Request("cboState"))
			session("freight_email") 		= Trim(Request("cboEmail"))
			session("freight_status") 		= Trim(Request("cboStatus"))			
			session("freight_initial_page") = 1
	end select
end sub

sub displayFreight
	dim iRecordCount
	iRecordCount = 0
    dim strSortBy
	dim strSortItem
    dim strSQL
	dim strPageResultNumber
	dim strRecordPerPage
	dim intRecordCount
	dim strTodayDate
	dim strDays

	strTodayDate = FormatDateTime(Date())

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 50

	if session("freight_status") = "" then
		session("freight_status") = "1"
	end if
	
	strSQL = "SELECT * FROM yma_freight "
	strSQL = strSQL & " WHERE "	
	strSQL = strSQL & "		receiver_state LIKE '%" & session("freight_state") & "%' "
	
	if Lcase(session("UsrUserName")) = "katet" or Lcase(session("UsrUserName")) = "helens" or Lcase(session("UsrUserName")) = "jackt" then
		strSQL = strSQL & "	AND supplier = 'cope' "
	end if
	
	strSQL = strSQL & "		AND (email LIKE '%" & session("freight_search") & "%' "
	strSQL = strSQL & "				OR receiver_name LIKE '%" & session("freight_search") & "%' "
	strSQL = strSQL & "				OR item_ref LIKE '%" & session("freight_search") & "%' "
	strSQL = strSQL & "				OR description LIKE '%" & session("freight_search") & "%' "
	strSQL = strSQL & "				OR item_ref2 LIKE '%" & session("freight_search") & "%' "
	strSQL = strSQL & "				OR description2 LIKE '%" & session("freight_search") & "%' "
	strSQL = strSQL & "				OR item_ref3 LIKE '%" & session("freight_search") & "%' "
	strSQL = strSQL & "				OR description3 LIKE '%" & session("freight_search") & "%' "
	strSQL = strSQL & "				OR item_ref4 LIKE '%" & session("freight_search") & "%' "
	strSQL = strSQL & "				OR description4 LIKE '%" & session("freight_search") & "%' "
	strSQL = strSQL & "				OR item_ref5 LIKE '%" & session("freight_search") & "%' "
	strSQL = strSQL & "				OR description5 LIKE '%" & session("freight_search") & "%' "
	strSQL = strSQL & "				OR item_ref6 LIKE '%" & session("freight_search") & "%' "
	strSQL = strSQL & "				OR description6 LIKE '%" & session("freight_search") & "%' "
	strSQL = strSQL & "				OR item_ref7 LIKE '%" & session("freight_search") & "%' "
	strSQL = strSQL & "				OR description7 LIKE '%" & session("freight_search") & "%' "
	strSQL = strSQL & "				OR item_ref8 LIKE '%" & session("freight_search") & "%' "
	strSQL = strSQL & "				OR description8 LIKE '%" & session("freight_search") & "%' "
	strSQL = strSQL & "				OR receiver_address LIKE '%" & session("freight_search") & "%' "
	strSQL = strSQL & "				OR items LIKE '%" & session("freight_search") & "%' "
    strSQL = strSQL & "             OR consignment_no LIKE '" & session("freight_search") & "%') "
	strSQL = strSQL & "		AND email LIKE '%" & session("freight_email") & "%' "
	strSQL = strSQL & "		AND status LIKE '%" & session("freight_status") & "%' "
	strSQL = strSQL & "	ORDER BY date_created DESC"
	
	'Response.Write strSQL & "<br>"

	rs.Open strSQL, conn

	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount

	Select Case Request("Action")
	    case "<<"
		    intpage = 1
			session("freight_initial_page") = intpage
	    case "<"
		    intpage = Request("intpage") - 1
			session("freight_initial_page") = intpage

			if session("freight_initial_page") < 1 then session("freight_initial_page") = 1
	    case ">"
		    intpage = Request("intpage") + 1
			session("freight_initial_page") = intpage

			if session("freight_initial_page") > intPageCount then session("freight_initial_page") = IntPageCount
	    Case ">>"
		    intpage = intPageCount
			session("freight_initial_page") = intpage
    end select

    strDisplayList = ""

	if not DB_RecSetIsEmpty(rs) Then

	    rs.AbsolutePage = session("freight_initial_page")

		For intRecord = 1 To rs.PageSize
			strRequester = LCase(rs("email"))
			'strRequester = Split(rs("email"),"@")

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

			strDisplayList = strDisplayList & "<td><a href=""update_freight.asp?id=" & rs("id") & """>" & rs("id") & "</a></td>"
			strDisplayList = strDisplayList & "<td>"
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
			strDisplayList = strDisplayList & "<td>" & rs("supplier") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("pickup_name") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("pickup_city") & " " & rs("pickup_postcode") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("pickup_date") & " - " & rs("pickup_time") & "</td>"
			if rs("return_pickup") = 1 then
				strDisplayList = strDisplayList & "<td bgcolor=""#FFFF00""><img src=images/tick.gif>"
			else
				strDisplayList = strDisplayList & "<td><img src=images/cross.gif>"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td><a href=""http://120.146.144.235/track.php?consignment=" & trim(rs("return_connote")) & """ target=""_blank"">" & trim(rs("return_connote")) & "</a></td>"
			strDisplayList = strDisplayList & "<td>" & rs("receiver_name") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("receiver_city") & " " & rs("receiver_postcode") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("delivery_date") & " - " & rs("delivery_time") & "</td>"
			strDisplayList = strDisplayList & "<td>"
			if rs("pickup") = 1 then
				strDisplayList = strDisplayList & "<img src=images/tick.gif>"
			else
				strDisplayList = strDisplayList & "<img src=images/cross.gif>"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td><a href=""http://120.146.144.235/track.php?consignment=" & trim(rs("consignment_no")) & """ target=""_blank"">" & trim(rs("consignment_no")) & "</a></td>"
			if rs("status") = 1 then
				strDisplayList = strDisplayList & "<td>Open</td>"
			else
				strDisplayList = strDisplayList & "<td class=""green_text"">Completed</td>"
			end if
			strDisplayList = strDisplayList & "<td><a onclick=""return confirm('Are you sure you want to delete " & rs("id") & " ?');"" href='delete_freight.asp?id=" & rs("id") & "'><img src=""images/btn_delete.gif"" border=""0""></a></td>"
			strDisplayList = strDisplayList & "</tr>"
			
			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next

	else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""17"" align=""center"">No freights found.</td></tr>"
	end if

	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""17"" class=""recordspaging"">"
	strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_freights.asp"" method=""post"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & session("freight_initial_page") & ">"

	if session("freight_initial_page") = 1 then
   		strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
	end if
	if session("freight_initial_page") = intpagecount or intRecordCount = 0 then
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value="">>"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value="">>"">"
	end if
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""txtSearch"" value=" & strSearch & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboDepartment"" value=" & strItemDepartment & ">"
	strDisplayList = strDisplayList & "<input type=""hidden"" name=""cboStatus"" value=" & strStatus & ">"
    strDisplayList = strDisplayList & "<br />"
    strDisplayList = strDisplayList & "Page: " & session("freight_initial_page") & " to " & intpagecount
	strDisplayList = strDisplayList & "<br />"
	strDisplayList = strDisplayList & "<h2>Search results: " & intRecordCount & " freights.</h2>"
    strDisplayList = strDisplayList & "</form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseDataBase()
end sub

sub main
	call UTL_validateLogin
	call setSearch

    if trim(session("freight_initial_page")) = "" then
    	session("freight_initial_page") = 1
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
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td><img src="images/icon_freight.jpg" border="0" alt="Freight Requests" /></td>
          <td valign="top"><img src="images/icon_excel.jpg" border="0" /> <a href="export_freight.asp?search=<%= session("freight_search") %>&state=<%= session("freight_state") %>&status=<%= session("freight_status") %>">Export</a></td>
          <td valign="top"><div class="alert alert-search">
              <form name="frmSearch" id="frmSearch" action="list_freights.asp?type=search" method="post" onsubmit="searchFreight()">
                <h3>Search Parameters:</h3>
                Requested by / Receiver name / Address / Items / Connote:
                <input type="text" name="txtSearch" size="25" value="<%= request("txtSearch") %>" maxlength="20" />
                <select name="cboEmail" onchange="searchFreight()">
                  <option <% if session("freight_email") = "" then Response.Write " selected" end if%> value="">All Requesters</option>
                  <option <% if session("freight_email") = "bernard" then Response.Write " selected" end if%> value="bernard">Bernard Crowe</option>
                  <option <% if session("freight_email") = "cameron" then Response.Write " selected" end if%> value="cameron">Cameron Tait</option>
                  <option <% if session("freight_email") = "herring" then Response.Write " selected" end if%> value="herring">Chris Herring</option>
                  <option <% if session("freight_email") = "moore" then Response.Write " selected" end if%> value="moore">Dale Moore</option>
                  <option <% if session("freight_email") = "henderson" then Response.Write " selected" end if%> value="henderson">Damien Henderson</option>
                  <option <% if session("freight_email") = "cooper" then Response.Write " selected" end if%> value="cooper">Daniel Cooper</option>
                  <option <% if session("freight_email") = "durante" then Response.Write " selected" end if%> value="durante">Dion Durante</option>
                  <option <% if session("freight_email") = "mcinnes" then Response.Write " selected" end if%> value="mcinnes">Euan McInnes</option>
                  <option <% if session("freight_email") = "felix" then Response.Write " selected" end if%> value="felix">Felix Elliot-Dedman</option>
                  <option <% if session("freight_email") = "grant" then Response.Write " selected" end if%> value="grant">Grant Lane</option>
                  <option <% if session("freight_email") = "jaclyn" then Response.Write " selected" end if%> value="jaclyn">Jaclyn Williams</option>
                  <option <% if session("freight_email") = "harvey" then Response.Write " selected" end if%> value="harvey">James Harvey</option>
                  <option <% if session("freight_email") = "jamie" then Response.Write " selected" end if%> value="jamie">Jamie Goff</option>
                  <option <% if session("freight_email") = "jason" then Response.Write " selected" end if%> value="jason">Jason Allen</option>
                  <option <% if session("freight_email") = "scholes" then Response.Write " selected" end if%> value="scholes">Johanna Scholes</option>
                  <option <% if session("freight_email") = "joseph" then Response.Write " selected" end if%> value="joseph">Joseph Pantalleresco</option>
                  <option <% if session("freight_email") = "justin" then Response.Write " selected" end if%> value="justin">Justin Doffay</option>
                  <option <% if session("freight_email") = "tietze" then Response.Write " selected" end if%> value="tietze">Kurt Tietze</option>
                  <option <% if session("freight_email") = "condon" then Response.Write " selected" end if%> value="condon">Mark Condon</option>
                  <option <% if session("freight_email") = "taylor" then Response.Write " selected" end if%> value="taylor">Mathew Taylor</option>
                  <option <% if session("freight_email") = "hughes" then Response.Write " selected" end if%> value="hughes">Mick Hughes</option>
                  <option <% if session("freight_email") = "biggin" then Response.Write " selected" end if%> value="biggin">Nathan Biggin</option>
                  <option <% if session("freight_email") = "shaun" then Response.Write " selected" end if%> value="shaun">Shaun McMahon</option>
                  <option <% if session("freight_email") = "vranch" then Response.Write " selected" end if%> value="vranch">Steven Vranch</option>
                  <option <% if session("freight_email") = "terry" then Response.Write " selected" end if%> value="terry">Terry McMahon</option>
                  <option <% if session("freight_email") = "steers" then Response.Write " selected" end if%> value="steers">Timothy Steers</option>
                  <option <% if session("freight_email") = "fischer" then Response.Write " selected" end if%> value="fischer">Wesley Fischer</option>
                </select>
                <select name="cboState" onchange="searchFreight()">
                  <%= strStateList %>
                </select>
                <select name="cboStatus" onchange="searchFreight()">
                  <option <% if session("freight_status") = "1" then Response.Write " selected" end if%> value="1">Open</option>
                  <option <% if session("freight_status") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
                </select>
                <input type="button" name="btnSearch" value="Search" onclick="searchFreight()" />
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
          <td>Requested by</td>
          <td>Supplier</td>
          <td>Pickup</td>
          <td>Location</td>
          <td>Date / time</td>
          <td bgcolor="#FF0000">Return to pickup?</td>
          <td>Return connote</td>
          <td bgcolor="#666666">Receiver</td>
          <td bgcolor="#666666">Location</td>
          <td bgcolor="#666666">Date / time</td>
          <td>Picked up?</td>
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