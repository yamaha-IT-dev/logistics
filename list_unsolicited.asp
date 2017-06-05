<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<% strSection = "unsolicited" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Incomplete Goods to Excel</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/javascript.js"></script>
<script type="text/javascript" src="include/jquery.js"></script>
<script type="text/javascript" src="include/main.js"></script>
<script language="JavaScript" type="text/javascript">
function searchUnsolicited(){
    var strSearch 		= document.forms[0].txtSearch.value;
	var strType  		= document.forms[0].cboType.value;
	var strDepartment  	= document.forms[0].cboDepartment.value;
	var strStatus 		= document.forms[0].cboStatus.value;

    document.location.href = 'list_unsolicited.asp?type=search&txtSearch=' + strSearch + '&cboType=' + strType + '&cboDepartment=' + strDepartment + '&cboStatus=' + strStatus;
}

function resetSearch(){
	document.location.href = 'list_unsolicited.asp?type=reset';
}
</script>
</head>
<body>
<%
sub setSearch
	select case Trim(Request("type"))
		case "reset"
			session("unsolicited_search") 		= ""
			session("unsolicited_instruction") 		= ""
			session("unsolicited_department") 	= ""
			session("unsolicited_status") 		= ""
			session("unsolicited_initial_page") = 1
		case "search"
			session("unsolicited_search") 		= trim(Request("txtSearch"))
			session("unsolicited_instruction") 		= request("cboType")
			session("unsolicited_department") 	= request("cboDepartment")
			session("unsolicited_status") 		= Trim(Request("cboStatus"))
			session("unsolicited_initial_page") = 1
	end select
end sub

sub displayUnsolicited
	dim iRecordCount
	iRecordCount = 0
    dim strSQL
	dim intRecordCount
	dim strTodayDate
	dim strDays

	strTodayDate = FormatDateTime(Date())

    call OpenDataBase()

	set rs = Server.CreateObject("ADODB.recordset")

	rs.CursorLocation = 3	'adUseClient
    rs.CursorType = 3		'adOpenStatic
	rs.PageSize = 100

	if session("unsolicited_status") = "" then
		session("unsolicited_status") = "1"
	end if	
	
	strSQL = "SELECT * FROM logistic_unsolicited "
	strSQL = strSQL & "	WHERE unsDepartment LIKE '%" & session("unsolicited_department") & "%' "	
	if session("unsolicited_instruction") <> "" then
		strSQL = strSQL & "		AND unsInstruction = '" & session("unsolicited_instruction") & "' "
	end if
	strSQL = strSQL & "		AND (unsItemCode LIKE '%" & session("unsolicited_search") & "%' "	
	strSQL = strSQL & "			OR unsDescription LIKE '%" & session("unsolicited_search") & "%' "
	strSQL = strSQL & "			OR unsConnote LIKE '%" & session("unsolicited_search") & "%' "
	strSQL = strSQL & "			OR unsGRA LIKE '%" & session("unsolicited_search") & "%' "
	strSQL = strSQL & "			OR unsDealer LIKE '%" & session("unsolicited_search") & "%' "	
	strSQL = strSQL & "			OR unsShipmentNo LIKE '%" & session("unsolicited_search") & "%' "
	strSQL = strSQL & "			OR unsComments LIKE '%" & session("unsolicited_search") & "%')"
	strSQL = strSQL & "		AND unsStatus LIKE '%" & session("unsolicited_status") & "%' "
	strSQL = strSQL & "	ORDER BY unsDateCreated DESC"

	'Response.Write strSQL & "<br>"

	rs.Open strSQL, conn

	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount

	Select Case Request("Action")
	    case "<<"
		    intpage = 1
			session("unsolicited_initial_page") = intpage
	    case "<"
		    intpage = Request("intpage") - 1
			session("unsolicited_initial_page") = intpage

			if session("unsolicited_initial_page") < 1 then session("unsolicited_initial_page") = 1
	    case ">"
		    intpage = Request("intpage") + 1
			session("unsolicited_initial_page") = intpage

			if session("unsolicited_initial_page") > intPageCount then session("unsolicited_initial_page") = IntPageCount
	    Case ">>"
		    intpage = intPageCount
			session("unsolicited_initial_page") = intpage
    end select

    strDisplayList = ""

	if not DB_RecSetIsEmpty(rs) Then

	    rs.AbsolutePage = session("unsolicited_initial_page")

		For intRecord = 1 To rs.PageSize
			strDays = DateDiff("d",rs("unsDateCreated"), strTodayDate)

			if (DateDiff("d",rs("unsDateModified"), strTodayDate) = 0) then
				if iRecordCount Mod 2 = 0 then
					strDisplayList = strDisplayList & "<tr class=""updated_today"">"
				else
					strDisplayList = strDisplayList & "<tr class=""updated_today_2"">"
				end if
			else
				'strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
				if iRecordCount Mod 2 = 0 then
					strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
				else
					strDisplayList = strDisplayList & "<tr class=""innerdoc_2"">"
				end if
			end if

			strDisplayList = strDisplayList & "<td nowrap><a href=""update_unsolicited.asp?id=" & rs("unsID") & """>" & rs("unsID") & "</a></td>"
			strDisplayList = strDisplayList & "<td>" & rs("unsCreatedBy") & " - " & WeekDayName(WeekDay(rs("unsDateCreated"))) & ", " & FormatDateTime(rs("unsDateCreated"),1) & "</td>"			
			strDisplayList = strDisplayList & "<td>" & rs("unsDepartment") & "</td>"			
			strDisplayList = strDisplayList & "<td><strong>" & Ucase(rs("unsItemCode")) & "</strong>"
			if DateDiff("d",rs("unsDateCreated"), strTodayDate) = 0 then
				strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=""0"">"
			end if
			strDisplayList = strDisplayList & "</td>"			
			strDisplayList = strDisplayList & "<td>" & rs("unsDescription") & "</td>"	
			strDisplayList = strDisplayList & "<td>" & rs("unsConnote") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("unsGRA") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("unsDealer") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("unsShipmentNo") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("unsQty") & "</td>"		
			strDisplayList = strDisplayList & "<td>"
			Select Case rs("unsInstruction")
				case 1
					strDisplayList = strDisplayList & "Move to 3XL"
				case 2
					strDisplayList = strDisplayList & "Move to 3S"
				case 3
					strDisplayList = strDisplayList & "Investigate"				
			end select
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("unsComments") & "</td>"
			if rs("unsStatus") = 1 then
				strDisplayList = strDisplayList & "<td class=""blue_text"">Open</td>"
			else
				strDisplayList = strDisplayList & "<td class=""green_text"">Completed</td>"
			end if
			strDisplayList = strDisplayList & "<td><a onclick=""return confirm('Are you sure you want to delete " & rs("unsItemCode") & " ?');"" href='delete_unsolicited.asp?id=" & rs("unsID") & "'><img src=""images/btn_delete.gif"" border=""0""></a></td>"
			strDisplayList = strDisplayList & "</tr>"

			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next
	else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""14"" align=""center"">No incomplete goods found.</td></tr>"
	end if
	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""14"" class=""recordspaging"">"
	strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_unsolicited.asp"" method=""post"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & session("unsolicited_initial_page") & ">"

	if session("unsolicited_initial_page") = 1 then
   		strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
	end if
	if session("unsolicited_initial_page") = intpagecount or intRecordCount = 0 then
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
    strDisplayList = strDisplayList & "Page: " & session("unsolicited_initial_page") & " to " & intpagecount
	strDisplayList = strDisplayList & "<br />"
	strDisplayList = strDisplayList & "<h2>Search results: " & intRecordCount & " incomplete goods.</h2>"
    strDisplayList = strDisplayList & "</form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseDataBase()
end sub

sub main
	call UTL_validateLogin
	call setSearch

    if trim(session("unsolicited_initial_page")) = "" then
    	session("unsolicited_initial_page") = 1
	end if

    call displayUnsolicited
end sub

call main

dim rs
dim intPageCount, intpage, intRecord
dim strDisplayList
%>
<table cellspacing="0" cellpadding="0" align="center" class="full_size_table" border="0">
  <!-- #include file="include/header.asp" -->
  <tr>
    <td class="first_content"><table cellpadding="5" cellspacing="0" border="0">
        <tr>
          <td valign="top"><img src="images/icon_return.jpg" border="0" alt="Warehouse Return" /></td>
          <td valign="top"><div class="alert alert-success"><img src="images/add_icon.png" border="0" align="bottom" /> <a href="add_unsolicited.asp">Add Incomplete Goods</a></div>           
            <p><img src="images/icon_excel.jpg" border="0" /> <a href="export_unsolicited.asp">Export</a></p>
            </td>
          <td valign="top"><div class="alert alert-search">
              <form name="frmSearch" id="frmSearch" action="list_unsolicited.asp?type=search" method="post" onsubmit="searchUnsolicited()">
                <h3>Search Parameters:</h3>
                Item / Shipment / Description / Connote / GRA :
                <input type="text" name="txtSearch" size="25" value="<%= request("txtSearch") %>" maxlength="20" />
                <select name="cboType" onchange="searchUnsolicited()">
                  <option value="">All Types</option>
                  <option <% if session("unsolicited_instruction") = "1" then Response.Write " selected" end if%> value="1">Move to 3XL</option>
                  <option <% if session("unsolicited_instruction") = "2" then Response.Write " selected" end if%> value="2">Move to 3S</option>
                  <option <% if session("unsolicited_instruction") = "3" then Response.Write " selected" end if%> value="3">Investigate</option>
                </select>
                <select name="cboDepartment" onchange="searchUnsolicited()">
                  <option value="">All Depts</option>
                  <option <% if session("unsolicited_department") = "AV" then Response.Write " selected" end if%> value="AV">AV</option>
                  <option <% if session("unsolicited_department") = "MPD" then Response.Write " selected" end if%> value="MPD">MPD</option>
                </select>
                <select name="cboStatus" onchange="searchUnsolicited()">
                  <option <% if session("unsolicited_status") = "1" then Response.Write " selected" end if%> value="1">Open</option>
                  <option <% if session("unsolicited_status") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
                </select>
                <input type="button" name="btnSearch" value="Search" onclick="searchUnsolicited()" />
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
          <td>Created</td>
          <td>Dept</td>
          <td>Item</td>
          <td>Description</td>
          <td>Connote</td>
          <td>GRA</td>
          <td>Dealer</td>
          <td>Shipment</td>
          <td>Qty</td>
          <td>Instruction</td>
          <td>Comments</td>
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