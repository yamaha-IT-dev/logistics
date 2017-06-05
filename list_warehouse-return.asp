<% response.cookies("current_URL_cookie_logistics") = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.Querystring %>
<!--#include file="include/connection_it.asp " -->
<% strSection = "quarantine" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Warehouse Return</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="include/stylesheet.css" type="text/css" />
<script type="text/javascript" src="include/javascript.js"></script>
<script type="text/javascript" src="include/jquery.js"></script>
<script type="text/javascript" src="include/main.js"></script>
<script>
function searchItem(){
    var strSearch 		= document.forms[0].txtSearch.value;
	var strType  		= document.forms[0].cboType.value;
	var strDepartment  	= document.forms[0].cboDepartment.value;
	var strInstruction  = document.forms[0].cboInstruction.value;
	var strStatus 		= document.forms[0].cboStatus.value;

    document.location.href = 'list_warehouse-return.asp?type=search&txtSearch=' + strSearch + '&cboType=' + strType + '&cboDepartment=' + strDepartment + '&cboInstruction=' + strInstruction + '&cboStatus=' + strStatus;
}

function resetSearch(){
	document.location.href = 'list_warehouse-return.asp?type=reset';
}
</script>
</head>
<body>
<%
sub setSearch
	select case Trim(Request("type"))
		case "reset"
			session("return_search") 		= ""
			session("return_type") 			= ""
			session("return_department") 	= ""
			session("return_instruction") 	= ""
			session("return_status") 		= ""
			session("return_initial_page") 	= 1
		case "search"
			session("return_search") 		= trim(Request("txtSearch"))
			session("return_type") 			= request("cboType")
			session("return_department") 	= request("cboDepartment")
			session("return_instruction") 	= request("cboInstruction")
			session("return_status") 		= Trim(Request("cboStatus"))
			session("return_initial_page") 	= 1
	end select
end sub

sub displayQuarantine
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

	if session("return_status") = "" then
		session("return_status") = "1"
	end if	
	
	strSQL = "SELECT * FROM yma_quarantines "
	strSQL = strSQL & "	WHERE department LIKE '%" & session("return_department") & "%' "
	
	if session("return_type") <> "" then
		strSQL = strSQL & "		AND return_type = '" & session("return_type") & "' "
	end if
	
	strSQL = strSQL & "		AND (item_code LIKE '%" & session("return_search") & "%' "
	strSQL = strSQL & "			OR shipment_no LIKE '%" & session("return_search") & "%' "
	strSQL = strSQL & "			OR description LIKE '%" & session("return_search") & "%' "
	strSQL = strSQL & "			OR return_carrier LIKE '%" & session("return_search") & "%' "
	strSQL = strSQL & "			OR return_connote LIKE '%" & session("return_search") & "%' "
	strSQL = strSQL & "			OR original_connote LIKE '%" & session("return_search") & "%' "
	strSQL = strSQL & "			OR gra LIKE '%" & session("return_search") & "%' "
	strSQL = strSQL & "			OR serial_no LIKE '%" & session("return_search") & "%' ) "
	strSQL = strSQL & "		AND instruction LIKE '%" & session("return_instruction") & "%' "
	strSQL = strSQL & "		AND status LIKE '%" & session("return_status") & "%' "
	strSQL = strSQL & "	ORDER BY date_created DESC"

	'Response.Write strSQL & "<br>"

	rs.Open strSQL, conn

	intPageCount = rs.PageCount
	intRecordCount = rs.recordcount

	Select Case Request("Action")
	    case "<<"
		    intpage = 1
			session("return_initial_page") = intpage
	    case "<"
		    intpage = Request("intpage") - 1
			session("return_initial_page") = intpage

			if session("return_initial_page") < 1 then session("return_initial_page") = 1
	    case ">"
		    intpage = Request("intpage") + 1
			session("return_initial_page") = intpage

			if session("return_initial_page") > intPageCount then session("return_initial_page") = IntPageCount
	    Case ">>"
		    intpage = intPageCount
			session("return_initial_page") = intpage
    end select

    strDisplayList = ""

	if not DB_RecSetIsEmpty(rs) Then

	    rs.AbsolutePage = session("return_initial_page")

		For intRecord = 1 To rs.PageSize
			strDays = DateDiff("d",rs("date_created"), strTodayDate)

			if (DateDiff("d",rs("date_modified"), strTodayDate) = 0) then
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

			strDisplayList = strDisplayList & "<td nowrap><a href=""update_warehouse-return.asp?id=" & rs("quarantine_id") & """>" & rs("quarantine_id") & "</a></td>"
			strDisplayList = strDisplayList & "<td>"
			Select Case rs("return_type")
				case 1
					strDisplayList = strDisplayList & "Managed"
				case 2
					strDisplayList = strDisplayList & "Un-addressed"
				case 0
					strDisplayList = strDisplayList & "Un-managed"				
				case else
			 		strDisplayList = strDisplayList & "-"
			end select
			strDisplayList = strDisplayList & "</td>"
			
			strDisplayList = strDisplayList & "<td>" & rs("department") & "</td>"
			strDisplayList = strDisplayList & "<td nowrap>" & strDays & " days</td>"
			strDisplayList = strDisplayList & "<td><strong>" & rs("item_code") & "</strong>"
			if DateDiff("d",rs("date_created"), strTodayDate) = 0 then
				strDisplayList = strDisplayList & " <img src=""images/icon_new.gif"" border=""0"">"
			end if
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>"
			Select Case rs("stock_type")
				case "1"
					strDisplayList = strDisplayList & "<font color=""red"">Damaged</font>"
				case "2"
					strDisplayList = strDisplayList & "<font color=""red"">Partial</font>"
				case else
			 		strDisplayList = strDisplayList & "-"
			end select
			strDisplayList = strDisplayList & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("return_connote") & "</td>"	
			strDisplayList = strDisplayList & "<td>" & rs("dealer") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("shipment_no") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("qty") & "</td>"
			strDisplayList = strDisplayList & "<td>"
			if rs("photos") = 1 then				
				strDisplayList = strDisplayList & "<a href=""file:\\YAMMAS22\quarantine\" & rs("quarantine_id") & """ target=""_blank"" class=""screenshot"" rel=""file:\\YAMMAS22\quarantine\" & rs("quarantine_id") & "\1.jpg""><img src=""images/camera_icon.gif"" border=""0""></a>"
			else
				strDisplayList = strDisplayList & "-"
			end if
			strDisplayList = strDisplayList & "</td>"

			strDisplayList = strDisplayList & "<td>" & rs("return_carrier") & "</td>"		
			strDisplayList = strDisplayList & "<td>" & rs("original_connote") & "</td>"
			strDisplayList = strDisplayList & "<td>" & rs("serial_no") & "</td>"
			
			strDisplayList = strDisplayList & "<td>"
			Select Case rs("reason_code")
				case "1"
					strDisplayList = strDisplayList & "Damaged in Transit"
				case "2"
					strDisplayList = strDisplayList & "Order Cancelled"
				case "3"
					strDisplayList = strDisplayList & "No longer required"
				case "4"
					strDisplayList = strDisplayList & "Order not in system"		
				case "5"
			 		strDisplayList = strDisplayList & "Other"
				case else
			 		strDisplayList = strDisplayList & "-"	
			end select
			strDisplayList = strDisplayList & "</td>"
			
			strDisplayList = strDisplayList & "<td>"
			Select Case rs("instruction")
				case "1"
					strDisplayList = strDisplayList & "Return to good stock 3T"
				case "2"
					strDisplayList = strDisplayList & "Transfer to Excel 3XL"
				case "3"
					strDisplayList = strDisplayList & "Resend to customer"
				case "4"
					strDisplayList = strDisplayList & "Damaged item to Excel - good stock to 3T"		
				case else
			 		strDisplayList = strDisplayList & "-"
			end select
			strDisplayList = strDisplayList & "</td>"
			
			strDisplayList = strDisplayList & "<td>" & rs("gra") & "</td>"
			'strDisplayList = strDisplayList & "<td>" & rs("comments") & "</td>"

			if rs("status") = 1 then
				strDisplayList = strDisplayList & "<td class=""blue_text"">Open</td>"
			else
				strDisplayList = strDisplayList & "<td class=""green_text"">Completed</td>"
			end if
			strDisplayList = strDisplayList & "<td><a onclick=""return confirm('Are you sure you want to delete " & rs("item_code") & " ?');"" href='delete_quarantine.asp?quarantine_id=" & rs("quarantine_id") & "'><img src=""images/btn_delete.gif"" border=""0""></a></td>"
			strDisplayList = strDisplayList & "</tr>"

			rs.movenext
			iRecordCount = iRecordCount + 1
			If rs.EOF Then Exit For
		next

	else
        strDisplayList = "<tr class=""innerdoc""><td colspan=""19"" align=""center"">No returns found.</td></tr>"
	end if

	strDisplayList = strDisplayList & "<tr class=""innerdoc"">"
	strDisplayList = strDisplayList & "<td colspan=""19"" class=""recordspaging"">"
	strDisplayList = strDisplayList & "<form name=""MovePage"" action=""list_warehouse-return.asp"" method=""post"">"
    strDisplayList = strDisplayList & "<input type=""hidden"" name=""intpage"" value=" & session("return_initial_page") & ">"

	if session("return_initial_page") = 1 then
   		strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input disabled type=""submit"" name=""action"" value=""<"">"
	else
		strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<<"">"
    	strDisplayList = strDisplayList & "<input type=""submit"" name=""action"" value=""<"">"
	end if
	if session("return_initial_page") = intpagecount or intRecordCount = 0 then
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
    strDisplayList = strDisplayList & "Page: " & session("return_initial_page") & " to " & intpagecount
	strDisplayList = strDisplayList & "<br />"
	strDisplayList = strDisplayList & "<h2>Search results: " & intRecordCount & " returns.</h2>"
    strDisplayList = strDisplayList & "</form>"
    strDisplayList = strDisplayList & "</td>"
    strDisplayList = strDisplayList & "</tr>"

    call CloseDataBase()
end sub

sub main
	call UTL_validateLogin
	call setSearch

    if trim(session("return_initial_page")) = "" then
    	session("return_initial_page") = 1
	end if

    call displayQuarantine
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
          <td valign="top"><div class="alert alert-success"><img src="images/add_icon.png" border="0" align="bottom" /> <a href="add_warehouse-return.asp">Add Warehouse Return</a></div>
            <% if Session("UsrLoginRole") = 1 then %>
            <p><img src="images/icon_excel.jpg" border="0" /> <a href="export_warehouse-return.asp">Export</a></p>
            <% end if %></td>
          <td valign="top"><div class="alert alert-search">
              <form name="frmSearch" id="frmSearch" action="list_warehouse-return.asp?type=search" method="post" onsubmit="searchItem()">
                <h3>Search Parameters:</h3>
                Item / Shipment / Description / Connotes / GRA / Serial no :
                <input type="text" name="txtSearch" size="25" value="<%= request("txtSearch") %>" maxlength="20" />
                <select name="cboType" onchange="searchItem()">
                  <option value="">All Types</option>
                  <option <% if session("return_type") = "1" then Response.Write " selected" end if%> value="1">Managed</option>
                  <option <% if session("return_type") = "0" then Response.Write " selected" end if%> value="0">Un-managed</option>
                  <option <% if session("return_type") = "2" then Response.Write " selected" end if%> value="2">Un-addressed</option>
                </select>
                <select name="cboDepartment" onchange="searchItem()">
                  <option value="">All Depts</option>
                  <option <% if session("return_department") = "AV" then Response.Write " selected" end if%> value="AV">AV</option>
                  <option <% if session("return_department") = "MPD" then Response.Write " selected" end if%> value="MPD">MPD</option>
                </select>
                <select name="cboInstruction" onchange="searchItem()">
                  <option value="">All Instructions</option>
                  <option <% if session("return_instruction") = "1" then Response.Write " selected" end if%> value="1">Return to good stock 3T</option>
                  <option <% if session("return_instruction") = "2" then Response.Write " selected" end if%> value="2">Transfer to Excel 3XL</option>
                  <option <% if session("return_instruction") = "3" then Response.Write " selected" end if%> value="3">Resend to customer</option>
                  <option <% if session("return_instruction") = "4" then Response.Write " selected" end if%> value="4">Damaged item to Excel - good stock to 3T</option>
                </select>
                <select name="cboStatus" onchange="searchItem()">
                  <option <% if session("return_status") = "1" then Response.Write " selected" end if%> value="1">Open</option>
                  <option <% if session("return_status") = "0" then Response.Write " selected" end if%> value="0">Completed</option>
                </select>
                <input type="button" name="btnSearch" value="Search" onclick="searchItem()" />
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
            <td>Type</td>
            <td>Dept</td>
            <td>Days</td>
            <td>Model no</td>
            <td>Stock</td>
            <td>Return connote</td>
            <td>Dealer</td>
            <td>Shipment</td>
            <td>Qty</td>
            <td>Photos</td>
            <td>Carrier</td>
            <td>Original connote</td>
            <td>Serial Number(s)</td>
            <td>Reason</td>
            <td>Instruction</td>
            <td>GRA</td>
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